VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visor de informes"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "frmVisReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   10545
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameCopia 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   4335
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   1920
         Min             =   1
         TabIndex        =   8
         Tag             =   "15000"
         Top             =   0
         Value           =   15000
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   3790
         TabIndex        =   5
         Top             =   75
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2820
         TabIndex        =   4
         Text            =   "1"
         Top             =   75
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Text            =   "1"
         Top             =   75
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   2220
         X2              =   2220
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   3340
         TabIndex        =   7
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   2300
         TabIndex        =   6
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Copias"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   480
      End
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      lastProp        =   600
      _cx             =   17595
      _cy             =   5318
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'COmentariio

Public Informe As String
'Public SubInformeConta As String 'SubInforme con conexion a la contabilidad. Conectar a las
                            'tablas de la BDatos correspondiente a la empresa: conta1, conta2, etc.
Public ConSubInforme As Boolean 'Si tiene subinforme ejecta la funcion AbrirSubInforme para enlazar esta a la BD correspondiente


'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean
Public Opcion As Integer
Public ExportarPDF As Boolean
Public EstaImpreso As Boolean

Public NumCopias As Integer ' (RAFA/ALZIRA 31082006) controla el n�mero de copias en un informe de impresion autom�tica

Public ForzarNombreImpresora As String 'Julio 2016. Albaranes TPV.  Sistema firma digital

Private WithEvents frmEx As frmVisReportExportar
Attribute frmEx.VB_VarHelpID = -1

Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim smrpt As CRAXDRT.Report

'Dim Argumentos() As String
Dim PrimeraVez As Boolean


Private Sub Command1_Click()
   
End Sub

Private Sub CRViewer1_ExportButtonClicked(UseDefault As Boolean)
    'Stop
    
    If vParamAplic.NumeroInstalacion = 4 Then
        
        UseDefault = False
        mrpt.ExportOptions.DiskFileName = ""
        Set frmEx = New frmVisReportExportar
        frmEx.Show vbModal
        Set frmEx = Nothing
        Screen.MousePointer = vbHourglass
        
        If mrpt.ExportOptions.DiskFileName <> "" Then
            
            
            
            Caption = String(200, " ") & "Documentos PDFs....... "
            Me.Refresh
            
            Screen.MousePointer = vbHourglass
            Me.MousePointer = vbHourglass
            DoEvents
            Screen.MousePointer = vbHourglass
            'Exportar
            mrpt.Export False


            If mrpt.ExportOptions.FormatType = crEFTPortableDocFormat Then
                If mrpt.PrintingStatus.Progress <> crPrintingCancelled Then
                    If mrpt.PrintingStatus.Progress = crPrintingCompleted Then HacerPDFSubDocumentos
                End If
            End If
            
            Me.MousePointer = vbDefault
            Caption = "Visor de informes"
        
        End If
        Screen.MousePointer = vbDefault
        
    End If
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
Dim Inicial As Integer

    On Error GoTo ePrintButtonClicked
        
    
      UseDefault = False
     
      If mrpt.PrinterSetupEx(Me.hwnd) = 0 Then
         
         'ok
         EstaImpreso = True
        
         
         If Text1(2).Text = "" Then
            mrpt.PrintOut False, CInt(Me.Text1(0).Text), , CInt(Val(Me.Text1(1).Text))
         Else
            mrpt.PrintOut False, CInt(Me.Text1(0).Text), , CInt(Val(Me.Text1(1).Text)), CInt(Val(Me.Text1(2).Text))
         End If
         
         If davidNumalbar > 0 Then DavidLogImpresionAlbaranes
     
     End If
    
    
    
    If EstaImpreso Then
        'Demomento solo EULER
        If vParamAplic.NumeroInstalacion = 4 Then
            If mrpt.PrintingStatus.Progress <> crPrintingCancelled Then
                If mrpt.PrintingStatus.Progress = crPrintingCompleted Then
                    'OK. Ha pulsadi imprimir
                    
                    
                    HacerImpresionSubDocumentos
                    
                End If
            End If
        End If

    End If
    
    
    
    Exit Sub
ePrintButtonClicked:
    MuestraError Err.Number, Err.Description
    
End Sub
Private Function PuedoCerrar(SegundoIncial As Single) As Boolean
Dim C As Integer
    PuedoCerrar = False
    If Not mrpt Is Nothing Then
        C = mrpt.PrintingStatus.Progress
        Debug.Print Now & " e:" & C
    Else
        C = 1
    End If
    
    If C = 2 Then
        DoEvents
        If Timer - SegundoIncial < 20 Then
            Screen.MousePointer = vbHourglass
            Espera 1
            'If Timer - SegundoIncial > 5 Then
        Else
            PuedoCerrar = True
        End If
    Else
        PuedoCerrar = True
    End If
End Function


Private Sub Form_Activate()
Dim Incio As Single
Dim fin As Boolean
    If PrimeraVez Then
    
    
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
           
        
            Screen.MousePointer = vbHourglass
            If SoloImprimir Then
                Incio = Timer
                Do
                    fin = PuedoCerrar(Incio)
                Loop Until fin
                Set mrpt = Nothing
                Set mapp = Nothing
            End If
            Unload Me
        Else
           ' PonerFocoBtn Text1(0)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim J As Integer
Dim NomImpre As String

    On Error GoTo Err_Carga

    
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    Screen.MousePointer = vbHourglass
    Set mapp = CreateObject("CrystalRuntime.Application")
    Set mrpt = mapp.OpenReport(Informe)
       
    If NumCopias = 0 Then NumCopias = 1
    Text1(0).Text = NumCopias
       
    'Conectar a la BD de la Empresa
    For I = 1 To mrpt.Database.Tables.Count
    
        'NUEVO 21 Mayo 2008
        'Puede que alguna tabla este vinculada a ARICONTA
        If LCase(CStr(mrpt.Database.Tables(I).ConnectionProperties.item("DSN"))) = "vconta" Then
            'A conta
            mrpt.Database.Tables(I).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
            'If (InStr(1, mrpt.Database.Tables(i).Name, "_") = 0) Then
            If RedireccionamosTabla(CStr(mrpt.Database.Tables(I).Name)) Then
               mrpt.Database.Tables(I).Location = "conta" & vParamAplic.NumeroConta & "." & mrpt.Database.Tables(I).Name
            End If
    
    
    
    
    
    
        ElseIf LCase(CStr(mrpt.Database.Tables(I).ConnectionProperties.item("DSN"))) = "mytelefono" Then
            'Detalle de llamada y poco mas
            mrpt.Database.Tables(I).SetLogOnInfo "mytelefono", , vParamAplic.UsuarioConta, vParamAplic.PasswordConta
        
        Else
            'A ariges
           mrpt.Database.Tables(I).SetLogOnInfo "vAriges", vEmpresa.BDAriges, vConfig.User, vConfig.password
    
           'If InStr(1, mrpt.Database.Tables(i).Name, "_") = 0 Then
           If RedireccionamosTabla(CStr(mrpt.Database.Tables(I).Name)) Then
                   mrpt.Database.Tables(I).Location = vEmpresa.BDAriges & "." & mrpt.Database.Tables(I).Name
           ElseIf InStr(1, mrpt.Database.Tables(I).Name, "alias") <> 0 Then
                J = InStr(1, mrpt.Database.Tables(I).Name, "_")
                mrpt.Database.Tables(I).Location = vEmpresa.BDAriges & "." & Mid(mrpt.Database.Tables(I).Name, 1, J - 1)
           End If
        End If
    Next I

'
'    If SubInformeConta <> "" Then
'        Set smrpt = mrpt.OpenSubreport(SubInformeConta)
'        For i = 1 To smrpt.Database.Tables.Count
'            smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
'            smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
'        Next i
'    End If
    
    'If ConSubInforme Then AbrirSubreport
    AbrirSubreportNuevo
    
    PrimeraVez = True
    
    CargaArgumentos
    
    
    mrpt.RecordSelectionFormula = FormulaSeleccion
'    mrpt.RecordSortFields

    If Opcion = 227 Then
    'Para INforme de Ventas por cliente
        If mrpt.FormulaFields.GetItemByName("pOrden").Text = "{tmpinformes.importe5}" Then
            mrpt.RecordSortFields.item(1).SortDirection = crDescendingOrder
        End If
    End If
    
    
    If ConSubInforme Then
        If Opcion = 228 Or Opcion = 240 Then
             smrpt.RecordSelectionFormula = mrpt.RecordSelectionFormula
        End If
    End If
    
    
    
    
'    If ConSubInforme Then
'        If Opcion = 50 Then
''            If Not (InStr(1, CStr(smrpt.RecordSelectionFormula), "tmpstockfec") > 0) Then
''                smrpt.RecordSelectionFormula = smrpt.RecordSelectionFormula & " and " & mrpt.RecordSelectionFormula
''            End If
'        End If
'    End If
    
    
    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If
    
     'lOS MARGENES
'    PonerMargen
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    
    
    EstaImpreso = False
'    mrpt.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
'    If Opcion = 93 Then 'TICKET
'        I = ObtenerTerminal
'        'Establecemos la impresora de ticket
'        NomImpre = NombreImpresoraTicket(I)
'
'        '## PRUEBAS
''        Dim X As Printer
''
'''        oImp = ObtenerImpresora(NomImpre)
''        For Each X In Printers
''           If X.DeviceName = NomImpre Then
''              ' La define como predeterminada del sistema.
''    '          Set Printer = X
''              ' Sale del bucle.
'''              ObtenerImpresora = X
''              Exit For
''           End If
''        Next
''        mrpt.SelectPrinter X.DriverName, X.DeviceName, X.Port
'        '##
'
'        mrpt.SelectPrinter "", NomImpre, ""
'    End If
    
    CRViewer1.ReportSource = mrpt
   
   
    If SoloImprimir Then
'        mrpt.PrinterName
        If ForzarNombreImpresora <> "" Then ForzarPonerNombreImpresora
        
        If NumCopias = 0 Then '(RAFA/ALZIRA 31082006) si se ha solicitado n�mero de copias se imprime ese n�mero
            mrpt.PrintOut False
        Else
            mrpt.PrintOut False, NumCopias
        End If
        EstaImpreso = True
        If davidNumalbar > 0 Then DavidLogImpresionAlbaranes
    Else
        CRViewer1.ViewReport
    End If
    

    
    
    
    Exit Sub
    
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical
    Set mapp = Nothing
    Set mrpt = Nothing
    Set smrpt = Nothing
End Sub



Private Sub ForzarPonerNombreImpresora()
    On Error GoTo eForzarPonerNombreImpresora
    
    mrpt.SelectPrinter "", ForzarNombreImpresora, ""
        
    Exit Sub
eForzarPonerNombreImpresora:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
        
    FrameCopia.Top = Me.CRViewer1.Top + 60
    FrameCopia.Left = CRViewer1.Width - CRViewer1.Left - 1600 - FrameCopia.Width
End Sub

Private Sub CargaArgumentos()
Dim Parametro As String
Dim I As Integer
    'El primer parametro es el nombre de la empresa para todas las empresas
    ' Por lo tanto concaatenaremos con otros parametros
    ' Y sumaremos uno
    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
    ' La cadena de parametros
    'Si esta asignaremos su valor
    
'    OtrosParametros = "|Emp= """ & vEmpresa.nomempre & """|" & OtrosParametros
Select Case NumeroParametros
Case 0
    '====Comenta: LAura
    'Solo se vacian los campos de formula que empiezan con "p" ya que estas
    'formulas se corresponden con paso de parametros al Report
    For I = 1 To mrpt.FormulaFields.Count
        If Left(Mid(mrpt.FormulaFields(I).Name, 3), 1) = "p" Then
            mrpt.FormulaFields(I).Text = """"""
        End If
    Next I
    '====
Case 1
    
    For I = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(I).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        'Debug.Print Parametro
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(I).Text = Parametro
        Else
'            mrpt.FormulaFields(I).Text = """"""
        End If
    Next I
    
Case Else
    NumeroParametros = NumeroParametros + 1
    
    For I = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(I).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(I).Text = Parametro
        End If
    Next I
'    mrpt.RecordSelectionFormula
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
    Set smrpt = Nothing
    NumCopias = 0 ' (RAFA/ALZIRA 31082006) por si acaso
    ForzarNombreImpresora = ""  'Para evitar problemas
    
    If vParamAplic.NumeroInstalacion = 4 Then ejecutar "DELETE from tmpImpresionAuxliar WHERE codusu = " & vUsu.codigo, False
    
End Sub


Private Function DevuelveValor(ByRef Valor As String) As Boolean
Dim I As Long
Dim J As Long

    Valor = "|" & Valor & "="
    DevuelveValor = False
    I = InStr(1, OtrosParametros, Valor, vbTextCompare)
    If I > 0 Then
        I = I + Len(Valor)
        J = InStr(I, OtrosParametros, "|")
        If J > 0 Then
            Valor = Mid(OtrosParametros, I, J - I)
            If Valor = "" Then
                Valor = " "
            Else
                If InStr(1, Valor, "chr(13)") = 0 Then CompruebaComillas Valor
            End If
            DevuelveValor = True
        End If
    End If
End Function


Private Sub CompruebaComillas(ByRef Valor1 As String)
Dim Aux As String
Dim J As Integer
Dim I As Integer

    If Mid(Valor1, 1, 1) = Chr(34) Then
        'Tiene comillas. Con lo cual tengo k poner las dobles
        Aux = Mid(Valor1, 2, Len(Valor1) - 2)
        I = -1
        Do
            J = I + 2
            I = InStr(J, Aux, """")
            If I > 0 Then
              Aux = Mid(Aux, 1, I - 1) & """" & Mid(Aux, I)
            End If
        Loop Until I = 0
        Aux = """" & Aux & """"
        Valor1 = Aux
    End If
End Sub

Private Sub Exportar()
    mrpt.ExportOptions.DiskFileName = App.Path & "\docum.pdf"
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    mrpt.ExportOptions.PDFExportAllPages = True
    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
    mrpt.Export False
    'Si ha generado bien entonces
    CadenaDesdeOtroForm = "OK"
End Sub

Private Sub PonerMargen()
Dim Cad As String
Dim I As Integer
    On Error GoTo EPon
    Cad = Dir(App.Path & "\*.mrg")
    If Cad <> "" Then
        I = InStr(1, Cad, ".")
        If I > 0 Then
            Cad = Mid(Cad, 1, I - 1)
            If IsNumeric(Cad) Then
                If Val(Cad) > 4000 Then Cad = "4000"
                If Val(Cad) > 0 Then
                    mrpt.BottomMargin = mrpt.BottomMargin + Val(Cad)
                End If
            End If
        End If
    End If
    
    Exit Sub
EPon:
    Err.Clear
End Sub






'FEBRERO 2012
'-----------------------------------------------------------------------------------
'Estoy teniendo problemas pq cuando "redirecciona" no lo hace para las tablas que llevan _
'
'   Que pasa, que las tablas de telefonia son tel_cab_ ...... y parece ser que no los direcciona
'
' Creon un abrirsub 2 donde para ver si redireccionamos la tabla , o no, lo hara una function
'======== LAURA
'Private Sub AbrirSubreport()
''Para cada subReport que encuentre en el Informe pone las tablas del subReport
''apuntando a la BD correspondiente
'Dim crxSection As CRAXDRT.Section
'Dim crxObject As Object
'Dim crxSubreportObject As CRAXDRT.SubreportObject
'Dim i As Byte
'
'    For Each crxSection In mrpt.Sections
'        For Each crxObject In crxSection.ReportObjects
'             If TypeOf crxObject Is SubreportObject Then
'                Set crxSubreportObject = crxObject
'                Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
'                For i = 1 To smrpt.Database.Tables.Count 'para cada tabla
'                    '------ A�ade Laura: 09/06/2005
'                    If smrpt.Database.Tables(i).ConnectionProperties.item("DSN") = "vAriges" Then
'                        smrpt.Database.Tables(i).SetLogOnInfo "vAriges", vEmpresa.BDAriges, vConfig.User, vConfig.password
'                        If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
'                        'If (InStr(1, smrpt.Database.Tables(i).Name, "_cmd") = 0) And (InStr(1, smrpt.Database.Tables(i).Name, "_alias") = 0) Then
'                           smrpt.Database.Tables(i).Location = vEmpresa.BDAriges & "." & smrpt.Database.Tables(i).Name
'                        End If
'                    ElseIf smrpt.Database.Tables(i).ConnectionProperties.item("DSN") = "vConta" Then
'                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
'                        If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
'                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
'                        End If
'
'                    ElseIf LCase(CStr(smrpt.Database.Tables(i).ConnectionProperties.item("DSN"))) = "mytelefono" Then
'                        'Detalle de llamada y poco mas
'                        smrpt.Database.Tables(i).SetLogOnInfo "myTelefono", "telefono", vParamAplic.UsuarioConta, vParamAplic.PasswordConta
'                        smrpt.Database.Tables(i).Location = "telefono." & smrpt.Database.Tables(i).Name
'
'                    End If
'                    '------
'                Next i
'             End If
'        Next crxObject
'    Next crxSection
'
'    Set crxSubreportObject = Nothing
'End Sub
'


Private Sub AbrirSubreportNuevo()
Dim crxSection As CRAXDRT.Section
Dim crxObject As Object
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim I As Byte

    For Each crxSection In mrpt.Sections
        For Each crxObject In crxSection.ReportObjects
             If TypeOf crxObject Is SubreportObject Then
                Set crxSubreportObject = crxObject
                Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
                For I = 1 To smrpt.Database.Tables.Count 'para cada tabla
                    '------ A�ade Laura: 09/06/2005
                    If smrpt.Database.Tables(I).ConnectionProperties.item("DSN") = "vAriges" Then
                        smrpt.Database.Tables(I).SetLogOnInfo "vAriges", vEmpresa.BDAriges, vConfig.User, vConfig.password
                        'If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                        If RedireccionamosTabla(CStr(smrpt.Database.Tables(I).Name)) Then
                           smrpt.Database.Tables(I).Location = vEmpresa.BDAriges & "." & smrpt.Database.Tables(I).Name
                        End If
                    ElseIf smrpt.Database.Tables(I).ConnectionProperties.item("DSN") = "vConta" Then
                        smrpt.Database.Tables(I).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
                        'If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                        If RedireccionamosTabla(CStr(smrpt.Database.Tables(I).Name)) Then
                           smrpt.Database.Tables(I).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(I).Name
                        End If
                        
                    ElseIf LCase(CStr(smrpt.Database.Tables(I).ConnectionProperties.item("DSN"))) = "mytelefono" Then
                        'Detalle de llamada y poco mas
                        smrpt.Database.Tables(I).SetLogOnInfo "myTelefono", "telefono", vParamAplic.UsuarioConta, vParamAplic.PasswordConta
                        smrpt.Database.Tables(I).Location = "telefono." & smrpt.Database.Tables(I).Name
                    
                    End If
                    '------
                Next I
             End If
        Next crxObject
    Next crxSection
    
    Set crxSubreportObject = Nothing
End Sub


Private Function RedireccionamosTabla(tabla As String) As Boolean
    'If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
    If InStr(1, tabla, "_") = 0 Then
        RedireccionamosTabla = True
    Else
        If Mid(tabla, 1, 3) = "tel" Then
            'tablas telefonia
            RedireccionamosTabla = True
        Else
            'resto
            RedireccionamosTabla = False
        End If
    End If
    
    
End Function




Private Sub HacerImpresionSubDocumentos()
Dim RN As ADODB.Recordset
    DoEvents
    Screen.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    Espera 1.5
    Set RN = New ADODB.Recordset
    RN.Open "Select * from tmpImpresionAuxliar WHERE codusu = " & vUsu.codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        
        lanzaImpresionShellDirecta Me.hwnd, DBLet(RN!Fichero, "T")
        Screen.MousePointer = vbHourglass
        Me.Refresh
        Espera 0.5
        
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
    Screen.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub


'Sobre el PDF creado de la exportacion, concatenar, con el programa
'asda
' el resto de archivos seleccionados
Private Sub HacerPDFSubDocumentos()
Dim RN As ADODB.Recordset
Dim Cad As String
Dim J As Integer


    DoEvents
    
    
    
    
    Screen.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    Espera 0.1
    Set RN = New ADODB.Recordset
    RN.Open "Select * from tmpImpresionAuxliar WHERE codusu = " & vUsu.codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    If Not RN.EOF Then
        If Dir(App.Path & "\temp\*.pdf", vbArchive) <> "" Then Kill App.Path & "\temp\*.pdf"
        FileCopy mrpt.ExportOptions.DiskFileName, App.Path & "\temp\tmp" & J & ".pdf"
        Kill mrpt.ExportOptions.DiskFileName
        
        While Not RN.EOF
            J = J + 1
            'Concatenamos
            Screen.MousePointer = vbHourglass
            
            Cad = """" & App.Path & "\temp\tmp" & J - 1 & ".pdf" & """" & " """ & RN!Fichero & """"
            
            Cad = """" & App.Path & "\pdftk.exe"" " & Cad & " cat output """ & App.Path & "\temp\tmp" & J & ".pdf" & """ verbose"
    
            Shell Cad, vbNormalFocus
            
            
            
            Me.Refresh
            Screen.MousePointer = vbHourglass
            Espera 0.5
            
            RN.MoveNext
        Wend
        RN.Close
        
        
        'Puede tardar bastante en crearse
        Cad = App.Path & "\temp\tmp" & J & ".pdf"
        J = 1
        Do
            If Dir(Cad, vbArchive) = "" Then
                Screen.MousePointer = vbHourglass
                DoEvents
                Espera 0.8
                J = J + 1
            Else
                Espera 0.1
                If CopiarFichero(Cad) Then
                    J = 35
                Else
                    J = 34
                End If
            End If
        Loop Until J > 30
        
        If J < 35 Then
            If J = 34 Then
                'YA ha dado msgb de error
            Else
                MsgBox "Error creando fichero: " & mrpt.ExportOptions.DiskFileName & vbCrLf & "Tiempo espera excedido", vbExclamation
            End If
        Else
            MsgBox "Fichero : " & mrpt.ExportOptions.DiskFileName & " creado con exito", vbInformation
        End If
    End If 'EOF
    Set RN = Nothing
    Screen.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub


Private Function CopiarFichero(Fichero As String) As Boolean
    On Error Resume Next
    FileCopy Fichero, mrpt.ExportOptions.DiskFileName
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        CopiarFichero = False
    Else
        CopiarFichero = True
    End If
End Function

Private Sub frmEx_DatoSeleccionado(CadenaSeleccion As String)
Dim Tipo As Byte

    Tipo = CByte(Mid(CadenaSeleccion, 1, 1))
    mrpt.ExportOptions.DiskFileName = Mid(CadenaSeleccion, 3)
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    
    'Fijamos valores
    If Tipo = 2 Then
        'WORD
        mrpt.ExportOptions.WORDWExportAllPages = True
        mrpt.ExportOptions.FormatType = crEFTWordForWindows
        'mrpt.ExportOptions.WORDWFirstPageNumber
    ElseIf Tipo = 1 Then
        mrpt.ExportOptions.FormatType = crEFTExcel97
        'mrpt.ExportOptions.ExcelFirstPageNumber

        mrpt.ExportOptions.ExcelExportAllPages = True
        mrpt.ExportOptions.ExcelTabHasColumnHeadings = False
        mrpt.ExportOptions.ExcelPageBreaks = False
        mrpt.ExportOptions.ExcelAreaType = crDetail
    Else
        mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
        mrpt.ExportOptions.PDFExportAllPages = True
        
    End If
End Sub



Private Sub Text1_GotFocus(Index As Integer)
     ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    'Si pulsa ESC
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Resetear As Boolean

    Text1(Index).Text = Trim(Text1(Index).Text)
    Resetear = False
    If Not PonerFormatoEntero(Text1(Index)) Then
        Resetear = True
        
    Else
        Text1(Index).Text = Abs(Text1(Index).Text) 'por si acaso
        If Index = 2 Then
            
        Else
            'NUmero de copias / Pagina inicio
            If Val(Text1(Index).Text) = 0 Then Resetear = True
        End If
    End If
    If Resetear Then
        If Index = 2 Then
            Text1(Index).Text = ""
        ElseIf Index = 1 Then
            Text1(Index).Text = "1"
        Else
            
            VScroll1.Value = 15000
            VScroll1.Tag = 15000
            Text1(0).Text = NumCopias
        End If
    Else
        'OK. Veamos que pagina final NO es mayor que inicio
        If Text1(2).Text <> "" Then
            If Val(Val(Me.Text1(1).Text)) > Val(Val(Me.Text1(2).Text)) Then Me.Text1(2).Text = Me.Text1(1).Text
        End If
    End If
End Sub

Private Sub SubirBajar(mas As Boolean)
Dim I As Integer
    
    If Not IsNumeric(Text1(0).Text) Then
        I = 1
    Else
        I = CInt(Val(Text1(0).Text))
    End If
    If mas Then
        I = I + 1
    Else
        I = I - 1
        If I < 1 Then I = 1
    End If
    Text1(0).Text = I
End Sub

Private Sub UpDown1_DownClick()
    SubirBajar False
End Sub

Private Sub UpDown1_UpClick()
    SubirBajar True
End Sub

Private Sub VScroll1_Change()
Dim Diferencia As Integer
    Diferencia = VScroll1.Tag - VScroll1.Value
    VScroll1.Tag = VScroll1.Value
    If Diferencia < 0 Then
        SubirBajar False
    Else
    
        SubirBajar True
    End If
End Sub
