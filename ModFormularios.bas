Attribute VB_Name = "ModFormularios"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'============================================================
'====== FUNCIONES GENERALES  ================================

'======== Añade: Laura

'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo ESugerirCodigo

    'SQL = "Select Max(codtipar) from stipar"
    SQL = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        SQL = SQL & " WHERE " & CondLineas
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If IsNumeric(RS.Fields(0)) Then
                SQL = CStr(RS.Fields(0) + 1)
            Else
                If Asc(Left(RS.Fields(0), 1)) <> 122 Then 'Z
                SQL = Left(RS.Fields(0), 1) & CStr(Asc(Right(RS.Fields(0), 1)) + 1)
                End If
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
    SugerirCodigoSiguienteStr = SQL
ESugerirCodigo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Public Function EsCodigoCero(Cod As String, Formato As String) As Boolean
'comprueba que para algunas tablas en las que el codigo 0000 se reserva para
'un valor genérico no se modifique ni se borre

    EsCodigoCero = False
    If Cod <> "" Then
        If Val(Cod) = Val(0) Then
            EsCodigoCero = True
            MsgBox "El código " & Formato & " no se puede modificar/eliminar.", vbExclamation
            Screen.MousePointer = vbDefault
        End If
    End If
End Function


Public Sub BloquearText1(ByRef Formulario As Form, Modo As Byte)
'Bloquea controles q se llamen TEXT1 si no estamos en Modo: 3.-Insertar, 4.-Modificar
'si estamos en modo modificar bloquea solo los campos que son clave primaria
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim i As Byte
Dim B As Boolean
Dim vtag As cTag
On Error Resume Next

    With Formulario
        B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'And ModoLineas = 1))
        
        For i = 0 To .Text1.Count - 1 'En principio todos los TExt1 tiene TAG
            Set vtag = New cTag
            vtag.Cargar .Text1(i)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 2 Or Modo = 4 Or Modo = 5) Then
                    .Text1(i).Locked = True
                    .Text1(i).BackColor = &H80000018 'amarillo claro
                Else
                    .Text1(i).Locked = Not B  '((Not b) And (Modo <> 1))
                    If B Then
                        .Text1(i).BackColor = vbWhite
                    Else
                        .Text1(i).BackColor = &H80000018 'amarillo claro
                    End If
                    If Modo = 3 Then .Text1(i).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            Else
                .Text1(i).Locked = Not B  '((Not b) And (Modo <> 1))
                If B Then
                    .Text1(i).BackColor = vbWhite
                Else
                    .Text1(i).BackColor = &H80000018 'amarillo claro
                End If
            End If
        Set vtag = Nothing
        Next i
        
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearTxt(ByRef Text As TextBox, B As Boolean, Optional EsContador As Boolean)
'Bloquea un control de tipo TextBox
'Si lo bloquea lo pone de color amarillo claro sino lo pone en color blanco (sino es contador)
'pero si es contador lo pone color azul claro
On Error Resume Next

    Text.Locked = B
    If Not B And Text.Enabled = False Then Text.Enabled = True
    If B Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
'            Text.BackColor = &H80000013 'Azul Claro
            Text.BackColor = &HFFFFC0   'Azul claro con vista
        Else
            Text.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Text.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearImg(ByRef imgF As Image, B As Boolean)
On Error Resume Next

    imgF.Enabled = Not B
    imgF.visible = Not B
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearCmb(ByRef cmb As ComboBox, B As Boolean, Optional EsContador As Boolean)
'Bloqueja un control de tipo ComboBox
'Si el bloqueja el posa de color gris claro, sino el posa de color blanc (sino es contador)
'pero si es contador el posa color blau clar
    On Error Resume Next

    cmb.Locked = B
    cmb.Enabled = True
    
    'cmb.Enabled = Not b
    
    'If Not b And Cmb.Enabled = False Then Cmb.Enabled = True
    If B Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
            cmb.BackColor = &H80000013 'Azul Claro
        Else
            cmb.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        cmb.BackColor = vbWhite
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub BloquearChecks(ByRef Formulario As Form, Modo As Byte)
'Bloquea controles  CheckBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim B As Boolean
Dim Control As Control
On Error Resume Next

    B = (Modo = 3 Or Modo = 4 Or Modo = 1)
    With Formulario
        For Each Control In Formulario.Controls
            If TypeOf Control Is CheckBox Then
                If Control.Name <> "chkVistaPrevia" Then
                    'modo Insertar o modificar
                    If Modo = 3 Or Modo = 4 Then
                        If Control.Value = 2 Then Control.Value = 1
                    End If
                    'modo consulta
                    If Modo = 0 Or Modo = 2 Then
                        If Control.Value = 1 Then Control.Value = 2
                    End If
                    'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                    If (Modo = 3 Or Modo = 1) Then Control.ListIndex = -1
                End If
            End If
        Next Control
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PonerLongCamposGnral(ByRef Formulario As Form, Modo As Byte, Opcion As Byte)
    Dim i As Integer
    
    On Error Resume Next

    With Formulario
        If Modo = 1 Then 'BUSQUEDA
            Select Case Opcion
                Case 1 'Para los TEXT1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = (.HelpContextID * 2) + 1 'el doble + 1
                            End If
                        End With
                    Next i
                
                Case 3 'para los TXTAUX
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = (.HelpContextID * 2) + 1 'el doble + 1
                            End If
                        End With
                    Next i
            End Select
            
        Else 'resto de modos
            Select Case Opcion
                Case 1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
                Case 3
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
            End Select
        End If
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub
 

Public Sub CargarICO(btn As CommandButton, Nombre As String)
    On Error Resume Next
    btn.Picture = LoadPicture(App.Path & "\iconos\" & Nombre)
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub DesplazamientoData(ByRef vData As Adodc, Index As Integer, Optional EsNuevo As Boolean)
'Para desplazarse por los registros de control Data
    If vData.Recordset.EOF Then Exit Sub
    If EsNuevo Then Index = Index - 1
    Select Case Index
        Case 0 'Primer Registro
            If Not vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 1 'Anterior
            vData.Recordset.MovePrevious
            If vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 2 'Siguiente
            vData.Recordset.MoveNext
            If vData.Recordset.EOF Then vData.Recordset.MoveLast
        Case 3 'Ultimo
            vData.Recordset.MoveLast
    End Select
End Sub




'===========================
Public Function SituarData(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registo que cumple vwhere
'para cuando la clave primaria esta formada por 1 campo
On Error GoTo ESituarData
        'Actualizamos el recordset
        vData.Refresh
        vData.Recordset.MoveFirst
        'El sql para que se situe en el registro en especial es el siguiente
        vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarData = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        If vData.Recordset.RecordCount > 0 Then vData.Recordset.MoveFirst
        SituarData = False
End Function




'===========================
Public Function SituarDataPosicion(ByRef vData As Adodc, NumPos As Long, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registro que ocupa la posicion NumPos
Dim TotalReg As Long
On Error GoTo ESituarDataPosicion
        'Actualizamos el recordset
'        vData.Refresh  'Refresh al cargar el grid

        TotalReg = vData.Recordset.RecordCount
        If vData.Recordset.EOF Then GoTo ESituarDataPosicion
        If NumPos <= TotalReg Then
            vData.Recordset.Move NumPos - 1
        Else
'            vData.Recordset.Move NumPos
            vData.Recordset.MoveLast
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataPosicion = True
        Exit Function
ESituarDataPosicion:
        If Err.Number <> 0 Then Err.Clear
        SituarDataPosicion = False
End Function


Public Sub DataLabelIndicador(ByRef vData As Adodc, ByRef lblIndicador As Label)

On Error Resume Next

        lblIndicador.Caption = ""
        lblIndicador.Caption = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
                
        If Err.Number <> 0 Then Err.Clear

End Sub
'===========================
Public Function SituarDataMULTI(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registo que cumple vwhere
On Error GoTo ESituarData
        'Actualizamos el recordset
        vData.Refresh
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find vData.Recordset, vWhere
        'vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataMULTI = False
End Function



'===========================
Public Function SituarRSetMULTI(ByRef vData As ADODB.Recordset, vWhere As String) As Boolean
'Situa un ADODB.Recordset en el registo que cumple vwhere
On Error GoTo ESituarData
    
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find2 vData, vWhere
        If vData.EOF Or vData.BOF Then GoTo ESituarData
        
        SituarRSetMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarRSetMULTI = False
End Function



Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)
'para el situarDataMULTI
On Error Resume Next
    Dim clone_rs As ADODB.Recordset
    Set clone_rs = oRs.Clone
    
    clone_rs.Filter = sCriteria
    
    If clone_rs.EOF Or clone_rs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    clone_rs.Close
    Set clone_rs = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub Multi_Find2(ByRef oRs As ADODB.Recordset, sCriteria As String)
'para el situarDataMULTI
On Error Resume Next

    oRs.Filter = ""
    oRs.MoveFirst
    oRs.Filter = sCriteria
    
    If oRs.EOF Or oRs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
'        x = oRs.AbsolutePosition
'     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg, Optional NoActualiza As Boolean) As Boolean
'NumReg: numero de registro que acabo de eliminar
'NoActualiza: si se hace el refresh o no, por defecto siempre se hace el refresh
'             pero si hemos eliminado de un Grid ya se hizo en el cargaGrid y
'             no lo volvemos a hacer para mantener las columnas bien.

    On Error GoTo ESituarDataElim
    
        If NoActualiza = False Then vData.Refresh
        
        If Not vData.Recordset.EOF Then    'Solo habia un registro
            If NumReg > vData.Recordset.RecordCount Then
                vData.Recordset.MoveLast
            Else
                vData.Recordset.MoveFirst
                vData.Recordset.Move NumReg - 1
            End If
            SituarDataTrasEliminar = True
        Else
            SituarDataTrasEliminar = False
        End If
        
ESituarDataElim:
    If Err.Number <> 0 Then
        Err.Clear
        SituarDataTrasEliminar = False
    End If
End Function


Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoBtn(ByRef btn As CommandButton)
On Error Resume Next
    If btn.visible Then btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoChk(ByRef chk As CheckBox)
On Error Resume Next
    chk.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoCbo(ByRef CBO As ComboBox)
On Error Resume Next
    CBO.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoGrid(ByRef DGrid As DataGrid)
    On Error Resume Next
    DGrid.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoOBj(ByRef obj As Variant)
    On Error Resume Next
    obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte, Optional cadkey As Integer)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If Modo = 5 Then Exit Sub
    
    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then
            Text.BackColor = vbLightBlue 'vbYellow  'Modo 1: Busqueda
        Else
            If Text.Locked Then 'si el control esta bloqueado pasamos el foco al sig. campo
                Text.BackColor = &H80000018 'amarillo claro
                 If cadkey = 0 Then cadkey = 40
                 KEYdown cadkey
                 Exit Sub
            Else
                Text.BackColor = vbWhite
            End If
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub ConseguirFocoLin(ByRef Text As TextBox, Optional cadkey As Integer)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

    With Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    'si el control esta bloqueado pasamos el foco al sig. campo
    If Text.Locked Then
        Text.BackColor = &H80000018 'amarillo claro
        If cadkey = 0 Then cadkey = 40
        KEYdown cadkey
'        Exit Sub

    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Function ObtenerCadKey(actCampo As Integer, sigCampo As Integer) As Integer
    Dim cadkey As Integer

    On Error Resume Next
    
    If actCampo > sigCampo Then
        cadkey = 38 'flecha superior
    Else
        cadkey = 40 'flecha inferior
    End If
    If sigCampo = 0 Then cadkey = 0
    
    ObtenerCadKey = cadkey
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Sub ConseguirfocoChk(Modo As Byte)
     If Modo = 0 Or Modo = 2 Then
        KEYpressGnral 13, Modo, False
    End If
End Sub


Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean
Dim Comprobar As Boolean
On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnral = False
        Exit Function
    End If
    
    With Text
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        If .BackColor = vbLightBlue Then .BackColor = vbWhite
        
        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (Modo <> 3 And Modo <> 4 And Modo <> 1) Then
            PerderFocoGnral = False
            Exit Function
        End If
        
        If Modo = 1 Then
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnral = False
                Exit Function
            End If
        End If
        PerderFocoGnral = True
    End With
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PerderFocoGnralLineas(ByRef txt As TextBox, ModoLineas As Byte) As Boolean
'Para el LostFocus de los txtAux de Mto de lineas
Dim Comprobar As Boolean

    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnralLineas = False
        Exit Function
    End If

    With txt
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        If .BackColor = vbYellow Then .BackColor = vbWhite
        
        'Si no estamos en modo: 1=Insertar o 2=Modificar , no hacer ninguna comprobacion
        If (ModoLineas <> 1 And ModoLineas <> 2) Then
            PerderFocoGnralLineas = False
            Exit Function
        End If
        
        If ModoLineas = 4 Then 'Busqueda
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnralLineas = False
                Exit Function
            End If
        End If
        
        'si el campo esta bloqueado no actualizar campos
        If .Locked Then
            PerderFocoGnralLineas = False
            Exit Function
        End If
        
        PerderFocoGnralLineas = True
    End With
    If Err.Number <> 0 Then Err.Clear
End Function


Public Sub AnyadirLinea(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
On Error Resume Next
    vDataGrid.AllowAddNew = True
    If vData.Recordset.RecordCount > 0 Then
        vDataGrid.HoldFields
        vData.Recordset.MoveLast
        vDataGrid.Row = vDataGrid.Row + 1
    End If
    
    vDataGrid.Enabled = False
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub DeseleccionaGrid(ByRef vDataGrid As DataGrid)
    On Error GoTo EDeseleccionaGrid

    While vDataGrid.SelBookmarks.Count > 0
        vDataGrid.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
    Err.Clear
End Sub

'Para forzar rowheiht
' vHeight
Public Sub CargaGridGnral(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, SQL As String, PrimeraVez As Boolean, Optional vHeight As Integer)
On Error GoTo ECargaGrid

    vDataGrid.Enabled = True
    
    'ANTES MARAZO 2011
    vData.ConnectionString = conn
    vData.RecordSource = SQL
    vData.CursorType = adOpenDynamic
    vData.LockType = adLockPessimistic
    vDataGrid.ScrollBars = dbgNone
    vData.Refresh
    
    Set vDataGrid.DataSource = vData
    vDataGrid.AllowRowSizing = False
    If vHeight = 0 Then vHeight = 290
    vDataGrid.RowHeight = vHeight

    If PrimeraVez Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaGrid", Err.Description
End Sub






Public Function DGrid_CambiarFila(ByRef DGrid As DataGrid) As Boolean
'comprobaciones para ejecutar el rowcolchange del datagrid
    If DGrid.Enabled = False Then Exit Function
    If DGrid.AllowAddNew = True Then Exit Function
    If Not DGrid.Columns(0).Text <> "" Then Exit Function
    
    DGrid_CambiarFila = True
End Function


Public Sub CargarCombo_SiNo(ByRef CBO As ComboBox)
'Carga un combo con los valores de opcion SI/NO
    On Error GoTo ErrCarga
    
    CBO.Clear
    
    CBO.AddItem "NO"
    CBO.ItemData(CBO.NewIndex) = 0
    
    CBO.AddItem "SI"
    CBO.ItemData(CBO.NewIndex) = 1
    
    Exit Sub
    
ErrCarga:
    MuestraError Err.Number, "Cargar combo.", Err.Description
End Sub


Public Sub CargarCombo_Tabla(ByRef CBO As ComboBox, NomTabla As String, NomCodigo As String, nomDescrip As String, Optional strWhere As String, Optional ItemNulo As Boolean, Optional Ordenacion As String)
'Carga un objeto ComboBox con los registros de una Tabla
'(IN) cbo: ComboBox en el q se van a cargar los datos
'(IN) nomTabla: nombre de la tabla de la q leeremos los datos a cargar
'(IN) nomCodigo: nombre del campo codigo de la tabla q queremos cargar
'(IN) nomDescrip: nombre del campo descripcion de la tabla a cargar
'(IN) strWhere: para filtrar los registros de la tabla q queremos cargar
'(IN) ItemNulo: si es true se añade el primer item con linea en blanco
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCombo
    
    CBO.Clear
    
    SQL = "SELECT " & NomCodigo & "," & nomDescrip & " FROM " & NomTabla
    If strWhere <> "" Then SQL = SQL & " WHERE " & strWhere
    SQL = SQL & " ORDER BY "
    If Ordenacion <> "" Then
        SQL = SQL & Ordenacion
    Else
        SQL = SQL & nomDescrip
    End If
    
'    If AbrirRecordset(SQL, RS) Then
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    '- si valor del parametro ItemNulo=true hay que añadir linea en blanco
    If Not RS.EOF And ItemNulo Then
        CBO.AddItem "  "
        CBO.ItemData(CBO.NewIndex) = 0
    End If
    
    If Not RS.EOF Then
        If IsNumeric(RS.Fields(0).Value) Then
            '- si el codigo NomCodigo es numerico en el ItemData se carga el campo clave primaria
            '- y en List la descripcion NomDescrip
            While Not RS.EOF
              CBO.AddItem RS.Fields(1).Value 'descrip
              CBO.ItemData(CBO.NewIndex) = RS.Fields(0).Value 'codigo
              RS.MoveNext
            Wend
        Else
            '- si el codigo NomCodigo en alfanumerico no se puede cargar
            '- el codigo en ItemData y cargamos un indice ficticio
            '- y en el List el campo codigo NomCodigo
            i = 1
            While Not RS.EOF
              CBO.AddItem RS.Fields(0).Value 'campo del codigo
              CBO.ItemData(CBO.NewIndex) = i
              i = i + 1
              RS.MoveNext
            Wend
        End If
    End If
'    End If
    
'    CerrarRecordset RS
    RS.Close
    Set RS = Nothing
    Exit Sub
    
ErrCombo:
    MuestraError Err.Number, "Cargar combo." & NomTabla, Err.Description
End Sub




Public Sub CargarCombo_TipMov(ByRef CBO As ComboBox, NomTabla As String, NomCodigo As String, nomDescrip As String, Optional strWhere As String, Optional ItemNulo As Boolean)
'Carga un objeto ComboBox con los registros de una Tabla
'(IN) cbo: ComboBox en el q se van a cargar los datos
'(IN) nomTabla: nombre de la tabla de la q leeremos los datos a cargar
'(IN) nomCodigo: nombre del campo codigo de la tabla q queremos cargar
'(IN) nomDescrip: nombre del campo descripcion de la tabla q queremos cargar
'(IN) strWhere: para filtrar los registros de la tabla q queremos cargar
'(IN) ItemNulo: si es true se añade el primer item con linea en blanco
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCombo
    
    CBO.Clear
    
    SQL = "SELECT " & NomCodigo & "," & nomDescrip & " FROM " & NomTabla
    If strWhere <> "" Then SQL = SQL & " WHERE " & strWhere
    SQL = SQL & " ORDER BY " & NomCodigo
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    '- si valor del parametro ItemNulo=true hay que añadir linea en blanco
    If Not RS.EOF And ItemNulo Then
        CBO.AddItem "  "
        CBO.ItemData(CBO.NewIndex) = 0
    End If
       
    i = 1
    While Not RS.EOF
        SQL = Replace(RS.Fields(1).Value, "Factura", "Fac.")
        SQL = RS.Fields(0).Value & " - " & SQL
        CBO.AddItem SQL 'campo del codigo
        CBO.ItemData(CBO.NewIndex) = i
        i = i + 1
        RS.MoveNext
    Wend

    RS.Close
    Set RS = Nothing
    Exit Sub
    
ErrCombo:
    MuestraError Err.Number, "Cargar combo." & NomTabla, Err.Description
End Sub



Public Sub CancelaADODC(ByRef vData As Adodc)
On Error Resume Next
    vData.Recordset.Cancel
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function EsVacio(ByRef campo As TextBox) As Boolean
    If (campo.Text = "" Or campo.Text = "0") Then
        EsVacio = True
    Else
        EsVacio = False
    End If
End Function


Public Sub DesplazamientoVisible(ByRef toolb As Toolbar, iniBoton As Byte, bol As Boolean, nreg As Byte)
'Oculta o Muestra las botones de  flechas de desplazamiento de la toolbar
Dim i As Byte

    Select Case nreg
        Case 0, 1 '0 o 1 registro no mostrar los botones despl.
            For i = iniBoton To iniBoton + 3
                toolb.Buttons(i).visible = False
            Next i
        Case Else '>1 reg, mostrar si bol
            For i = iniBoton To iniBoton + 3
                toolb.Buttons(i).visible = bol
            Next i
    End Select
End Sub


Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte)
'Pone el titulo del label lblIndicador
    lblIndicador.FontBold = True
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
            
        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
            
            
            
        
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub


Public Sub ActualizarToolbarGnral(ByRef Toolbar1 As Toolbar, Modo As Byte, Kmodo As Byte, posic As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner
Dim B As Boolean
    
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    '---- [2012] DAVID : Materias activas y equivalencias
    B = (Modo = 5 Or Modo = 6 Or Modo = 7 Or Modo = 8 Or Modo = 9 Or Modo = 10)
    
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    ' 19/12/2011  Modo 9=  Materias activas
    If (B) And (Kmodo <> 5 And Kmodo <> 6 And Kmodo <> 7 And Kmodo <> 8 And Kmodo <> 9 And Kmodo <> 10) Then  'Cabecera
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(posic).Image = 3
        Toolbar1.Buttons(posic).ToolTipText = "Nuevo"
        '-- Modificar
        Toolbar1.Buttons(posic + 1).Image = 4
        Toolbar1.Buttons(posic + 1).ToolTipText = "Modificar"
        '-- eliminar
        Toolbar1.Buttons(posic + 2).Image = 5
        Toolbar1.Buttons(posic + 2).ToolTipText = "Eliminar"
    End If
    
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN (se añade modo 8)
    If (Kmodo = 5 Or Kmodo = 6 Or Kmodo = 7 Or Kmodo = 8 Or Kmodo = 9 Or Kmodo = 10) Then 'Lineas
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(posic).Image = 12
        Toolbar1.Buttons(posic).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(posic + 1).Image = 13
        Toolbar1.Buttons(posic + 1).ToolTipText = "Modificar linea"
        '-- eliminar
        Toolbar1.Buttons(posic + 2).Image = 14
        Toolbar1.Buttons(posic + 2).ToolTipText = "Eliminar linea"
    End If
End Sub


Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then cerrar = True
    End If
End Sub


Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            SendKeys "+{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub KEYdownLineas(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 37 'Desplazamiento Flecha Izquierda
            SendKeys "+{tab}"
        Case 38 'Desplazamieto Flecha Hacia Arriba
            SendKeys "+{tab}"
        Case 39 'Desplaz. Flecha Derecha
            SendKeys "{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub SituarMultiTextFinal(ByRef txt As TextBox)
    On Error GoTo ErrMulti
    
    'situa el cursor del text multilinea al final para poder empezar a escribir
    If txt.Text <> "" And txt.MultiLine And txt.Enabled And txt.Locked = False Then SendKeys "^{END}"
    Exit Sub
     
ErrMulti:
    MuestraError Err.Number, "", Err.Description
End Sub




Public Sub SituarCombo(ByRef CBO As ComboBox, Valor As Long)
Dim i As Byte

    On Error Resume Next

        For i = 0 To CBO.ListCount - 1
            If CBO.ItemData(i) = Val(Valor) Then
                CBO.ListIndex = i
                Exit For
            End If
        Next i
        If i = CBO.ListCount Then CBO.ListIndex = -1
    
    If Err.Number <> 0 Then
        CBO.ListIndex = -1
        Err.Clear
    End If
End Sub


Public Function ObtenerAlto(ByRef vDataGrid As DataGrid, Optional alto As Integer) As Single
Dim anc As Single
    anc = vDataGrid.Top + alto
    If vDataGrid.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + vDataGrid.RowTop(vDataGrid.Row)
    End If
    ObtenerAlto = anc
End Function


'*********** LAURA : 13/09/2005
Public Function EsEnteroNew(texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEnteroNew = False

    If Not IsNumeric(texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 0 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 0 Then res = False
        End If
    End If
    EsEnteroNew = res
End Function




'=================================
'******** DAVID (NO LA USO)
Public Function EsEntero(texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function



Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As cTag
Dim Cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New cTag
    mTag.Cargar T
    If mTag.Cargado Then
       Cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEnteroNew(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & Cad & " tiene que ser un número entero.", vbExclamation
        PonerFoco T
    Else
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function


'******* IMPORTANTE
' El tipo de datos CURRENCY solo admite 4 decimales
Public Function PonerFormatoDecimal_Single(ByRef T As TextBox, tipoF As Single) As Boolean
Dim Valor2 As Single
Dim PEntera As Currency
Dim NoOK As Boolean
Dim Tg As cTag
Dim FormatoTag As String
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(10,4)
'  3 -> Decimal(10,2)    '¡FORMATO CANTIDAD
'  4 -> Decimal(4,2)
'  5 -> Decimal(8,4)
'  6 -> Decimal(8,2)
'  7 -> Decimal(5,2)
'
'
'  8 -> Lo que ponga en su TAG
'  9 ->  Formato precio2. Para cuando podamos parametrizarlo


    PonerFormatoDecimal_Single = False
    If T.Text = "" Then Exit Function
    NoOK = False
    With T
        If Not EsNumerico(.Text) Then
'            .Text = ""
            PonerFoco T
        Else
            If InStr(1, .Text, ",") > 0 Then
                Valor = ImporteFormateadoSingle(.Text)
            Else
                Valor = CSng(TransformaPuntosComas(.Text))
            End If

            'Comprobar la longitud de la Parte Entera
            PEntera = Int(Valor)
            Select Case tipoF 'Comprobar longitud
                Case 1 'Decimal(12,2)
                    If Len(PEntera) > 10 Then
                        MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 2, 9 'Decimal(10,4)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 3 'Decimal(10,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 4 'Decimal(4,2)
                    If Len(CStr(PEntera)) > 2 Then
                        MsgBox "El valor no puede ser mayor de 99,99", vbExclamation
                        NoOK = True
                    End If
                Case 5 'Decimal(8,4)
                    If Len(CStr(PEntera)) > 4 Then
                        MsgBox "El valor no puede ser mayor de 9999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 6 'Decimal(8,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 7 'Decimal(5,2)
                    '---- Laura: 05/10/2006
                    '# ANTES:   If Len(CStr(PEntera)) > 3 Then
                    If Len(CStr(Abs(PEntera))) > 3 Then
                    '----
                        MsgBox "El valor no puede ser mayor de 100,00", vbExclamation
                        NoOK = True
                    End If
                    
                Case 8
                    'David 12 Feb 07
                    'Lo que ponga en su tag
                    Set Tg = New cTag
                    If Not Tg.Cargar(T) Then NoOK = True
                    FormatoTag = Tg.Formato
                    Set Tg = Nothing
            End Select
            
            If NoOK Then
                .Text = ""
                T.SetFocus
                PonerFormatoDecimal_Single = False
                Exit Function
            Else
                PonerFormatoDecimal_Single = True
            End If

            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(10,4)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoCantidad)
                Case 4 'Formato Decimal(4,2)
                    .Text = Format(Valor, FormatoDescuento)
                Case 5 'Formato Decimal(8,4)
                    .Text = Format(Valor, FormatoKms)
                Case 6 'Formato Decimal(8,2)
                    .Text = Format(Valor, FormatoCantidad2)
                Case 7 'Formato Decimal(5,2)
                    .Text = Format(Valor, "##0.00")
                Case 8
                    .Text = Format(Valor, FormatoTag)
                Case 9 'Formato Decimal(10,5). Intentaremos parametrizarlo
                    .Text = Format(Valor, FormatoPrecio2)
                    
            End Select
        End If
    End With

End Function




'******* IMPORTANTE
'Leer el procedimiento de arriba.   IMPORTANTE:   PonerFormatoDecimal_Single
'---------------------------------------------------------------------------------
Public Function PonerFormatoDecimal(ByRef T As TextBox, tipoF As Single) As Boolean
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(10,4)
'  3 -> Decimal(10,2)    '¡FORMATO CANTIDAD
'  4 -> Decimal(4,2)
'  5 -> Decimal(8,4)
'  6 -> Decimal(8,2)
'  7 -> Decimal(5,2)
'
'
'  8 -> Lo que ponga en su TAG
Dim Valor As Currency
Dim PEntera As Currency
Dim NoOK As Boolean
Dim Tg As cTag
Dim FormatoTag As String

On Error GoTo ePonerFormatoDecimal


    PonerFormatoDecimal = False
    If T.Text = "" Then Exit Function
    NoOK = False
    With T
        If Not EsNumerico(.Text) Then
'            .Text = ""
            PonerFoco T
        Else
            If InStr(1, .Text, ",") > 0 Then
                Valor = ImporteFormateado(.Text)
            Else
                Valor = CCur(TransformaPuntosComas(.Text))
            End If

            'Comprobar la longitud de la Parte Entera
            PEntera = Int(Valor)
            Select Case tipoF 'Comprobar longitud
                Case 1 'Decimal(12,2)
                    If Len(PEntera) > 10 Then
                        MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 2 'Decimal(10,4)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 3 'Decimal(10,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 4 'Decimal(4,2)
                    If Len(CStr(Abs(PEntera))) > 2 Then
                        MsgBox "El valor no puede ser mayor de 99,99", vbExclamation
                        NoOK = True
                    End If
                Case 5 'Decimal(8,4)
                    If Len(CStr(PEntera)) > 4 Then
                        MsgBox "El valor no puede ser mayor de 9999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 6 'Decimal(8,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 7 'Decimal(5,2)
                    '---- Laura: 05/10/2006
                    '# ANTES:   If Len(CStr(PEntera)) > 3 Then
                    If Len(CStr(Abs(PEntera))) > 3 Then
                    '----
                        MsgBox "El valor no puede ser mayor de 100,00", vbExclamation
                        NoOK = True
                    End If
                    
                Case 8
                    'David 12 Feb 07
                    'Lo que ponga en su tag
                    Set Tg = New cTag
                    If Not Tg.Cargar(T) Then NoOK = True
                    FormatoTag = Tg.Formato
                    Set Tg = Nothing
            End Select
            
            If NoOK Then
                .Text = ""
                T.SetFocus
                PonerFormatoDecimal = False
                Exit Function
            Else
                PonerFormatoDecimal = True
            End If

            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(10,4)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoCantidad)
                Case 4 'Formato Decimal(4,2)
                    .Text = Format(Valor, FormatoDescuento)
                Case 5 'Formato Decimal(8,4)
                    .Text = Format(Valor, FormatoKms)
                Case 6 'Formato Decimal(8,2)
                    .Text = Format(Valor, FormatoCantidad2)
                Case 7 'Formato Decimal(5,2)
                    .Text = Format(Valor, "##0.00")
                Case 8
                    .Text = Format(Valor, FormatoTag)
            End Select
        End If
    End With
    
    Exit Function
    
ePonerFormatoDecimal:
    MuestraError Err.Number
    T.Text = ""
    
End Function


Public Function PonerNombreDeCod(ByRef txt As TextBox, BD As Byte, tabla As String, campo As String, Optional Codigo As String, Optional texto As String, Optional Tipo As String) As String
'Devuelve el nombre/Descripción asociado al Código correspondiente
'Además pone formato al campo txt del código a partir del Tag
Dim SQL As String
Dim devuelve As String
Dim vtag As cTag
Dim ValorCodigo As String
On Error GoTo EPonerNombresDeCod

    ValorCodigo = txt.Text
    If ValorCodigo <> "" Then
        Set vtag = New cTag
        If vtag.Cargar(txt) Then
            If Codigo = "" Then Codigo = vtag.columna
            If Tipo = "" Then Tipo = vtag.TipoDato
            SQL = DevuelveDesdeBD(BD, campo, tabla, Codigo, ValorCodigo, Tipo)
            If vtag.TipoDato = "N" Then ValorCodigo = Format(ValorCodigo, vtag.Formato)
            If SQL = "" Then
                If texto = "" Then
                    devuelve = "No existe " & vtag.Nombre & ": " & ValorCodigo
                Else
                    devuelve = "No existe " & texto & ": " & ValorCodigo
                End If
                MsgBox devuelve, vbExclamation
'                Txt.Text = ""
                'si ponemos foco bucle
'                PonerFoco Txt
'                Txt.SetFocus
            Else
                PonerNombreDeCod = SQL 'Descripcion del codigo
                'Poner valor codigo formateado
                txt.Text = ValorCodigo 'Valor codigo formateado
            End If
        End If
        Set vtag = Nothing
    Else
        PonerNombreDeCod = ""
    End If
    Exit Function
EPonerNombresDeCod:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Nombre asociado a código: " & Codigo, Err.Description
End Function


Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vtag As cTag
Dim devuelve As String
On Error GoTo EExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vtag = New cTag
            If vtag.Cargar(T) Then
                devuelve = DevuelveDesdeBDNew(conAri, vtag.tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                If devuelve <> "" Then
                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                    ExisteCP = True
                End If
            End If
            Set vtag = Nothing
        End If
    End If
EExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código", Err.Description
End Function



Public Sub SubirItemList(ByRef LView As ListView)
'Subir el item seleccionado del listview una posicion
Dim i As Byte, Item As Byte
Dim Aux As String
On Error Resume Next
   
    For i = 2 To LView.ListItems.Count
        If LView.ListItems(i).Selected Then
            Item = i
            Aux = LView.ListItems(i).Text
            LView.ListItems(i).Text = LView.ListItems(i - 1).Text
            LView.ListItems(i - 1).Text = Aux
        End If
    Next i
    If Item <> 0 Then
        LView.ListItems(Item).Selected = False
        LView.ListItems(Item - 1).Selected = True
    End If
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BajarItemList(ByRef LView As ListView)
'Bajar el item seleccionado del listview una posicion
Dim i As Byte, Item As Byte
Dim Aux As String
On Error Resume Next

    For i = 1 To LView.ListItems.Count - 1
        If LView.ListItems(i).Selected Then
            Item = i
            Aux = LView.ListItems(i).Text
            LView.ListItems(i).Text = LView.ListItems(i + 1).Text
            LView.ListItems(i + 1).Text = Aux
        End If
    Next i
    If Item <> 0 Then
        LView.ListItems(Item).Selected = False
        LView.ListItems(Item + 1).Selected = True
    End If
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub CargarProgres(ByRef PBar As ProgressBar, Valor As Integer)
On Error Resume Next
    PBar.Max = 100
    PBar.Value = 0
    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub IncrementarProgres(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub CargarProgresNew(ByRef PBar As ProgressBar, Valor As Long)
On Error Resume Next
    PBar.Value = 0
    
    If Valor > 32765 Then
        PBar.Max = 32765
    Else
        If Valor = 0 Then Valor = 1
        PBar.Max = CInt(Valor)
    End If
    
'    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    If PBar.Value < PBar.Max Then PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PosicionarComboDes(ByRef CBO As ComboBox, Valor As String)
Dim i As Byte

    For i = 0 To CBO.ListCount - 1
        If Trim(CBO.List(i)) = Trim(Valor) Then
            CBO.ListIndex = i
            Exit For
        End If
    Next i
    If i = CBO.ListCount Then CBO.ListIndex = -1
    
End Sub



Public Sub PosicionarCombo(ByRef Combo1 As ComboBox, Valor As Integer)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(J) = Valor Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub





'============================================================
'====== FUNCIONES PARA ARIGES  ==============================
'============================================================

Public Function PonerNombreCuenta(ByRef txt As TextBox, Modo As Byte, Optional clien As String) As String
Dim DevfrmCCtas As String
Dim SQL As String

     If txt.Text = "" Then
         PonerNombreCuenta = ""
         Exit Function
    End If
    DevfrmCCtas = txt.Text
    If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
        If InStr(SQL, "No existe la cuenta") > 0 Then
            txt.Text = DevfrmCCtas
            
            If (Modo = 3 Or Modo = 4) Then  'si insertar o modificar
                SQL = SQL & "  ¿Desea crearla?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                    'SI MODO es insetar NO me sirve el metodo anterior. Pq? Pq aun no he creado el cli/prov
                    'De momento pondre una marca en el texto de descripcion para que la cree
                    If Modo = 3 Then
                        PonerNombreCuenta = vbCrearNuevaCta
                                                
                    
                    
                    Else
                        If InStr(1, txt.Tag, "sclien") > 0 Then
                            InsertarCuentaCble DevfrmCCtas, clien
                        ElseIf InStr(1, txt.Tag, "sprove") > 0 Then
                            InsertarCuentaCble DevfrmCCtas, "", clien
                        ' ---- [02/10/2009] (LAURA): crear cuenta en familias articulos
                        ElseIf InStr(1, txt.Tag, "sfamia") > 0 Then
                            InsertarCuentaCble DevfrmCCtas, "", , clien
                        ' ----
                        End If
                        PonerNombreCuenta = clien
                    End If
                Else
                    'DAVID
                    'Si me dice que no quiere crearla, pongo el txt a blanco
                    txt.Text = ""
                End If
            Else
                MsgBox SQL, vbExclamation
            End If
        Else
            txt.Text = DevfrmCCtas
            PonerNombreCuenta = SQL
        End If
    Else
        If Modo = 3 Or Modo = 4 Or Modo = 1 Then 'si insertar o modificar
            MsgBox SQL, vbExclamation
'            PonerNombreCuenta = ""
        End If
'        Txt.Text = ""
        PonerNombreCuenta = ""
'        ConseguirFoco Txt, Modo
        PonerFoco txt
    End If
    DevfrmCCtas = ""
End Function

'He cambiado el metodo a public
Public Function InsertarCuentaCble(Cuenta As String, cadClien As String, Optional cadProve As String, Optional cadFamia As String) As Boolean
Dim SQL As String
Dim vClien As CCliente
Dim vProve As CProveedor
Dim Aux As String
Dim B As Boolean

    On Error GoTo EInsCta
    
    
    
    SQL = ""
    
    SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,forpa, ctabanco, "
    If vParamAplic.ContabilidadNueva Then
        SQL = SQL & "codpais,iban"
    Else
        SQL = SQL & "pais,entidad, oficina, CC, CuentaBa,iban"
    End If
    SQL = SQL & ") VALUES (" & DBSet(Cuenta, "T") & ","
    
    If cadClien <> "" Then
        Set vClien = New CCliente
        If vClien.LeerDatos(cadClien) Then
            SQL = SQL & DBSet(vClien.Nombre, "T") & ",'S',1," & DBSet(vClien.Nombre, "T") & "," & DBSet(vClien.Domicilio, "T") & ","
            SQL = SQL & DBSet(vClien.CPostal, "T") & "," & DBSet(vClien.Poblacion, "T") & "," & DBSet(vClien.Provincia, "T") & "," & DBSet(vClien.NIF, "T") & "," & DBSet(vClien.EMailAdm, "T") & "," & DBSet(vClien.WebClien, "T") & "," & ValorNulo
            'Forma pago y cuenta banco por defecto
            SQL = SQL & "," & DBSet(vClien.ForPago, "N", "S") & "," & ValorNulo & "," & DBSet(vClien.PAIS, "T", "S") & ","
            
            'PAIS
            If vParamAplic.ContabilidadNueva Then
                Aux = MiFormat(vClien.Iban, "") & MiFormat(vClien.Banco, "0000") & MiFormat(vClien.Sucursal, "0000") & MiFormat(vClien.DigControl, "00") & MiFormat(vClien.CuentaBan, "0000000000")
                SQL = SQL & DBSet(Aux, "T")

            Else
                
                If vClien.Banco = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.Banco, "0000") & "'"
                End If
                SQL = SQL & ","
                If vClien.Sucursal = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.Sucursal, "0000") & "'"
                End If
                SQL = SQL & ","
                If vClien.DigControl = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.DigControl, "00") & "'"
                End If
                SQL = SQL & ","
                If vClien.CuentaBan = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.CuentaBan, "0000000000") & "'"
                End If
                            
                SQL = SQL & ","
                If vClien.Iban = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & vClien.Iban & "'"
                End If
            End If
            SQL = SQL & ")"
            
            ConnConta.Execute SQL
            cadClien = vClien.Nombre
            B = True
        Else
            B = False
        End If
        Set vClien = Nothing
    End If
    
    If cadProve <> "" Then
        Set vProve = New CProveedor
        If vProve.LeerDatos(cadProve) Then
            SQL = SQL & DBSet(vProve.Nombre, "T") & ",'S',1," & DBSet(vProve.Nombre, "T") & "," & DBSet(vProve.Domicilio, "T") & ","
            SQL = SQL & DBSet(vProve.CPostal, "T") & "," & DBSet(vProve.Poblacion, "T") & "," & DBSet(vProve.Provincia, "T") & "," & DBSet(vProve.NIF, "T") & ","
            SQL = SQL & DBSet(vProve.EMailAdmon, "T") & "," & DBSet(vProve.WebProve, "T") & "," & ValorNulo
            'Forma pago y cuenta banco por defecto
            cadProve = DevuelveDesdeBD(conAri, "codmacta", "sbanpr", "codbanpr", vProve.BancoPropio)
            SQL = SQL & "," & DBSet(vProve.ForPago, "N", "S") & "," & DBSet(cadProve, "N", "S") & ","
            cadProve = ""
            
            'PAIS
            If vParamAplic.ContabilidadNueva Then
                Aux = DevuelveDesdeBD(conAri, "codpais", "sprove", "codprove", vProve.Codigo)
                If Aux = "" Then
                    Aux = ValorNulo
                Else
                    Aux = DBSet(Aux, "T")
                End If
            Else
                Aux = ValorNulo
            End If
            SQL = SQL & Aux & ","
            
            'CuentaBAnco
            If vParamAplic.ContabilidadNueva Then
                Aux = MiFormat(vProve.Iban, "") & MiFormat(vProve.Banco, "0000") & MiFormat(vProve.Sucursal, "0000") & MiFormat(vProve.DigControl, "00") & MiFormat(vProve.CuentaBan, "0000000000")
                SQL = SQL & DBSet(Aux, "T")

            Else
                
                If vProve.Banco = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.Banco, "0000") & "'"
                End If
                SQL = SQL & ","
                If vProve.Sucursal = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.Sucursal, "0000") & "'"
                End If
                SQL = SQL & ","
                If vProve.DigControl = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.DigControl, "00") & "'"
                End If
                SQL = SQL & ","
                If vProve.CuentaBan = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.CuentaBan, "0000000000") & "'"
                End If
                            
                SQL = SQL & ","
                If vProve.Iban = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & vProve.Iban & "'"
                End If
            End If
            SQL = SQL & ")"
            
            
            
            ConnConta.Execute SQL
            cadProve = vProve.Nombre
            B = True
        Else
            B = False
        End If
        Set vProve = Nothing
    
    ' ---- [02/10/2009] (LAURA): crear cuenta en familias articulos
    ElseIf cadFamia <> "" Then 'cuentas familias articulos
        SQL = SQL & DBSet(cadFamia, "T") & ",'S',0," & DBSet(cadFamia, "T") & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        'Forma pago y cuenta banco por defecto
        If Not vParamAplic.ContabilidadNueva Then SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ")"
        
        ConnConta.Execute SQL
        B = True
    ' ----
    End If
    
EInsCta:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Insertando cuenta contable", Err.Description
    End If
    InsertarCuentaCble = B
End Function




Public Function ModificarCtaContabilidad(esCliente As Boolean, Cuenta As String, Codigo As Long) As Boolean
Dim SQL As String
Dim vClien As CCliente
Dim vProve As CProveedor
Dim B As Boolean
Dim vL As String
Dim Aux As String

    On Error GoTo EModCta
    
    
    'Para el LOG
    If esCliente Then
        vL = "Cli: "
    Else
        vL = "Pro: "
    End If
    vL = vL & Format(Codigo, "0000") & " - " & Cuenta & "  "
      
'    SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,
'   codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,pais,forpa, ctabanco) "
'    SQL = SQL & " VALUES (" & DBSet(cuenta, "T") & ","
    SQL = "UPDATE cuentas SET "
    If esCliente Then
        Set vClien = New CCliente
        If vClien.LeerDatos(CStr(Codigo)) Then
            
            SQL = SQL & "nommacta=" & DBSet(vClien.Nombre, "T") & ", razosoci=" & DBSet(vClien.Nombre, "T") & ", dirdatos= " & DBSet(vClien.Domicilio, "T")
            SQL = SQL & ", codposta = " & DBSet(vClien.CPostal, "T") & ", despobla=" & DBSet(vClien.Poblacion, "T") & ", desprovi=" & DBSet(vClien.Provincia, "T")
            SQL = SQL & ", nifdatos=" & DBSet(vClien.NIF, "T") & ", maidatos=" & DBSet(vClien.EMailAdm, "T", "S") & ", webdatos=" & DBSet(vClien.WebClien, "T", "S")
            'Forma pago y cuenta banco por defecto
            SQL = SQL & ",forpa=" & DBSet(vClien.ForPago, "N", "S")
            
            'Cuenta bancaria
            If vParamAplic.ContabilidadNueva Then
                
                Aux = MiFormat(vClien.Iban, "") & MiFormat(vClien.Banco, "0000") & MiFormat(vClien.Sucursal, "0000") & MiFormat(vClien.DigControl, "00") & MiFormat(vClien.CuentaBan, "0000000000")
                SQL = SQL & ", iban=" & DBSet(Aux, "T")
            
            Else
                SQL = SQL & ", entidad="
                If vClien.Banco = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.Banco, "0000") & "'"
                End If
                SQL = SQL & ", oficina="
                If vClien.Sucursal = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.Sucursal, "0000") & "'"
                End If
                SQL = SQL & ", " & IIf(vParamAplic.ContabilidadNueva, "control", "CC") & " ="
                If vClien.DigControl = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.DigControl, "00") & "'"
                End If
                SQL = SQL & ", cuentaba="
                If vClien.CuentaBan = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vClien.CuentaBan, "0000000000") & "'"
                End If
                            
                SQL = SQL & ", iban="
                If vClien.Iban = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & vClien.Iban & "'"
                End If
            End If
            
            'Pais
            If vParamAplic.ContabilidadNueva Then
                
                Aux = DevuelveDesdeBD(conAri, "codpais", "sclien", "codclien", vClien.Codigo)
                If Aux = "" Then
                    Aux = ValorNulo
                Else
                    'Tiene pais. Grabaraemos:
                    '   Si es intracom
'                    If Aux = "ES" Then
'                        'Aux = "ESPAÑA"
'                    Else
'                        Aux = DevuelveDesdeBD(conConta, "concat(codpais,'|',codpais,'|',intracom,'|')", "paises", "codpais", Aux, "T")
'                        If Aux = "" Or Aux = "|||" Then
'                            Aux = ValorNulo
'                        Else
'                            If RecuperaValor(Aux, 3) = "0" Then
'                                'Extranjero
'                                Aux = RecuperaValor(Aux, 2) & " (" & RecuperaValor(Aux, 1) & ")"
'                            Else
'                                'Intracomunitaria
'                                Aux = RecuperaValor(Aux, 1) & " " & RecuperaValor(Aux, 2)
'                            End If
'                        End If
'                    End If
                
                End If
                If Aux <> ValorNulo Then Aux = DBSet(Aux, "T")
                SQL = SQL & ", codpais=" & Aux
            End If
            
            If vParamAplic.ContabilidadNueva Then SQL = SQL & " , maidatos=" & DBSet(vClien.EMailAdm, "T", "S")
                
            
            
            SQL = SQL & " WHERE codmacta = " & DBSet(Cuenta, "T")
            
            
            ConnConta.Execute SQL
            B = True
            
            
            
            vL = vL & vClien.Nombre
            
            
        Else
            B = False
        End If
        Set vClien = Nothing
    
    
    Else
        Set vProve = New CProveedor
        If vProve.LeerDatos(CStr(Codigo)) Then
            SQL = SQL & "nommacta=" & DBSet(vProve.Nombre, "T") & ", razosoci=" & DBSet(vProve.Nombre, "T") & ", dirdatos= " & DBSet(vProve.Domicilio, "T")
            SQL = SQL & ", codposta = " & DBSet(vProve.CPostal, "T") & ", despobla=" & DBSet(vProve.Poblacion, "T") & ",desprovi=" & DBSet(vProve.Provincia, "T")
            SQL = SQL & ", nifdatos=" & DBSet(vProve.NIF, "T") & ", maidatos=" & DBSet(vProve.EMailAdmon, "T", "S") & ", webdatos=" & DBSet(vProve.WebProve, "T", "S")
            'Forma pago y cuenta banco por defecto
            SQL = SQL & ",forpa=" & DBSet(vProve.ForPago, "N", "S")
            
            'Cuenta bancaria
            If vParamAplic.ContabilidadNueva Then
            
                 Aux = MiFormat(vProve.Iban, "") & MiFormat(vProve.Banco, "0000") & MiFormat(vProve.Sucursal, "0000") & MiFormat(vProve.DigControl, "00") & MiFormat(vProve.CuentaBan, "0000000000")
                SQL = SQL & ", iban=" & DBSet(Aux, "T")
            
            Else
                SQL = SQL & ", entidad="
                If vProve.Banco = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.Banco, "0000") & "'"
                End If
                SQL = SQL & ", oficina="
                If vProve.Sucursal = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.Sucursal, "0000") & "'"
                End If
                
                SQL = SQL & ", " & IIf(vParamAplic.ContabilidadNueva, "control", "CC") & " ="
                If vProve.DigControl = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.DigControl, "00") & "'"
                End If
                SQL = SQL & ", cuentaba="
                If vProve.CuentaBan = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & Format(vProve.CuentaBan, "0000000000") & "'"
                End If
                SQL = SQL & ", iban="
                If vProve.Iban = "" Then
                    SQL = SQL & "NULL"
                Else
                    SQL = SQL & "'" & vProve.Iban & "'"
                End If
                
            End If
            
            SQL = SQL & " WHERE codmacta = " & DBSet(Cuenta, "T")
            
            
            ConnConta.Execute SQL
            vL = vL & vProve.Nombre
            B = True
            
            
            
            
            
        Else
            B = False
        End If
        Set vProve = Nothing
    End If
    
    If B Then
        'METEMOS UN LOG
        
        Codigo = InStr(1, SQL, "dirdatos")
        If Codigo > 0 Then SQL = Mid(SQL, Codigo)
        
        Codigo = InStrRev(SQL, "WHERE codmacta")
        If Codigo > 0 Then SQL = Mid(SQL, 1, Codigo - 1)
        
        
        SQL = Replace(SQL, "dirdatos", "Dir")
        SQL = Replace(SQL, "codposta", "CP")
        SQL = Replace(SQL, "despobla", "Pob")
        SQL = Replace(SQL, "desprovi", "Prv")
        SQL = Replace(SQL, "nifdatos", "NIF")
        SQL = Replace(SQL, "entidad", "Banco")
        SQL = Replace(SQL, "oficina=", "")
        SQL = Replace(SQL, "CC=", "")
        SQL = Replace(SQL, "cuentaba=", "")
        SQL = Replace(SQL, "maidatos=", "@")
        SQL = Replace(SQL, "webdatos=", "W:")
        
        
        SQL = Replace(SQL, "'", "")
        'Quito las comas
        Codigo = InStr(1, SQL, "CP =")
        SQL = Replace(SQL, ",", "")
        
        vL = vL & vbCrLf & SQL
        Set LOG = New cLOG
        LOG.Insertar 18, vUsu, vL
        Set LOG = Nothing
            
        
    End If
    
    
    
EModCta:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Description, "Modificando cuenta contable", Err.Description
    End If
    ModificarCtaContabilidad = B
    
End Function


'Si es "" devuelve "" , si no, devuelve el campo formateado
Private Function MiFormat(Valor As String, Formato As String) As String
    If Trim(Valor) = "" Then
       MiFormat = ""
    Else
        If Formato = "" Then
            MiFormat = Valor
        Else
            MiFormat = Format(Valor, Formato)
        End If
    End If
End Function


'He cambiado el metodo a public
Public Function InsertarCuentaCbleDescripcion(Cuenta As String, Descripcion As String) As Boolean
Dim SQL As String


   
    
    SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci) "
    SQL = SQL & " VALUES ('" & Cuenta & "','" & DevNombreSQL(Descripcion) & "','S',0,'" & DevNombreSQL(Descripcion) & "')"
    ConnConta.Execute SQL

End Function






Public Function ComprobarHayStock(stockOrig As Single, stockTras As Single, codArtic As String, NomArtic As String, tipoMov As String)
'IN: stockOrig: stock existente en almacen Origen
'    stockTras: stock a traspasar del origen a otro almacen
Dim B As Boolean
Dim devuelve As String

    ComprobarHayStock = False
    If stockOrig >= CSng(stockTras) Then
    'Si cantidad en stock > cantidad a traspasar entonces
        B = True
    Else    'No hay suficiente stock en almacen origen
        devuelve = "Control de Stock : " & vbCrLf
        devuelve = devuelve & "---------------------- " & vbCrLf & vbCrLf
        devuelve = devuelve & " No hay suficiente Stock en el Almacen del Artículo:"
        devuelve = devuelve & vbCrLf & " Código:   " & codArtic & vbCrLf
        devuelve = devuelve & " Desc.: " & NomArtic & vbCrLf & vbCrLf
        devuelve = devuelve & "(Stock=" & stockOrig & ")"

        If tipoMov = "OFE" Then
            MsgBox devuelve, vbInformation
        Else
            If vParamAplic.ControlStock Then
            'Si hay control Stock no permitir traspaso
                B = False
                Select Case tipoMov
                    Case "REG"
                        devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Movimiento de Almacen. "
                    Case "TRA"
                        devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Traspaso de Almacen. "
                End Select
                MsgBox devuelve, vbExclamation
            Else
                Select Case tipoMov
                Case "REG"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Movimiento de Almacen? "
                Case "TRA"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Traspaso de Almacen? "
                End Select
                If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then
                    B = True
                Else
                    B = False
                End If
            End If
        End If
    End If
    ComprobarHayStock = B
End Function


Public Function LanzaHomeGnral(nomWeb As String) As Boolean
On Error GoTo ELanzaHome

    LanzaHomeGnral = False
    'Obtenemos la pagina web de los parametros
'    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", Opcion, "codigo", "1", "N")
'    If CadenaDesdeOtroForm = "" Then
'        MsgBox "Falta configurar los datos en Parámetros de la Aplicación.", vbExclamation
'        Exit Function
'    End If
    If nomWeb = "" Then
        MsgBox "No hay una dirección Web para mostrar.", vbInformation
        Exit Function
    End If

    'Lanzamos
'    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    If vConfig.Explorador <> "" Then
        Shell vConfig.Explorador & " " & nomWeb, vbMaximizedFocus
        LanzaHomeGnral = True
    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, nomWeb & vbCrLf & Err.Description
End Function


Public Function LanzaMailGnral(dirMail As String) As Boolean
'LLama al Programa de Correo (Outlook,...)
On Error GoTo ELanzaHome

    LanzaMailGnral = False
    If dirMail = "" Then
        MsgBox "No hay dirección e-mail a la que enviar.", vbExclamation
        Exit Function
    End If

    Call ShellExecute(hwnd, "Open", "mailto: " & dirMail, "", "", vbNormalFocus)
    LanzaMailGnral = True
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, vbCrLf & Err.Description
End Function




Public Function PonerArticuloEan(ByRef txtCod As TextBox, ByRef txtNom As TextBox, codAlm As String, tipoMov As String, Optional Modo As Byte, Optional AntCodArtic As String, Optional sConLotes As Boolean, Optional ByRef txtProv As String, Optional StatusArticuloMayorCero As Boolean) As Boolean
Dim C As String
    PonerArticuloEan = False
    '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN y quitar de la cabecera
'    C = DevuelveDesdeBD(conAri, "codartic", "sartic", "codigoea", txtCod.Text, "T")
    C = DevuelveDesdeBD(conAri, "codartic", "sarti3", "codigoea", txtCod.Text, "T")
    '----
    
    If C = "" Then
        MsgBox "El codigo EAN no corresponde con ningun articulo", vbExclamation
    Else
        txtCod.Text = C
        PonerArticuloEan = PonerArticulo(txtCod, txtNom, codAlm, tipoMov, Modo, AntCodArtic, sConLotes, txtProv, StatusArticuloMayorCero)
    End If
End Function


Public Function PonerArticulo(ByRef txtCod As TextBox, ByRef txtNom As TextBox, codAlm As String, tipoMov As String, Optional Modo As Byte, Optional AntCodArtic As String, Optional sConLotes As Boolean, Optional ByRef txtProv As String, Optional ByRef StatusMayor0 As Boolean) As Boolean
'Poner el codigo y nombre correcto de un Articulo
'IN: txtCod: codigo del articulo
'    txtNom: nombre del articulo
'    codAlm: codigo del almacen en el que comprobamos si se esta inventariando (almacen en el que se va a realizar el movimiento)
Dim vArtic As CArticulo
Dim Bloquea As Boolean
Dim NoSeguir As Boolean
Dim N As Integer
Dim L As Integer

    PonerArticulo = False
    sConLotes = False
    
    Set vArtic = New CArticulo
        
    If vParamAplic.DigitosCodartic > 0 Then
        'Sutituyen punto por cero hasta digitos codartic
        N = InStr(1, txtCod.Text, ".")
        If N > 0 Then
            L = vParamAplic.DigitosCodartic - (Len(txtCod.Text) - 1)
            If L > 0 Then txtCod.Text = Mid(txtCod.Text, 1, N - 1) & String(CLng(L), "0") & Mid(txtCod.Text, N + 1)
        End If
    End If
    If vArtic.Existe(txtCod.Text) Then
        If vArtic.LeerDatos(txtCod.Text) Then
            'comprobar que existe el articulo en el almacen del movimiento
            If vArtic.ExisteEnAlmacen(codAlm) Then
            
                'comprobar si el articulo esta inventariandose
                If tipoMov = "OFE" Then
                    NoSeguir = False 'Dejamos que continue. No bloqueamos por inventario
                Else
                    NoSeguir = vArtic.EnInventario(codAlm)
                End If
                If NoSeguir Then
                    If Modo = 1 Then 'Insertar lineas
                        txtCod.Text = ""
                        txtNom.Text = ""
                    End If
                    PonerFoco txtCod
                Else
                    'comprobar si el articulo esta bloqueado
                    vArtic.MostrarStatusArtic Bloquea
                    StatusMayor0 = vArtic.Status > 0
                    
                    If Bloquea Then 'El articulo esta bloqueado
                        If Modo = 1 Then
                            txtCod.Text = ""
                            txtNom.Text = ""
                        End If
                        PonerFoco txtCod
                    Else 'Articulo OK
                        PonerArticulo = True
                        
                        'Si es articulo DE VARIOS podemos modificar la descripción del articulo, sino bloqueamos.
                        If Not EsArticuloVarios(txtCod.Text) Then
                            BloquearTxt txtNom, True
                            'si insertando lineas
                            'If Modo = 1 Then txtNom.Text = vArtic.Nombre
                            txtNom.Text = vArtic.Nombre
                        Else
                            'si insertando lineas
                            If Modo = 1 Then
                                txtNom.Text = vArtic.Nombre
                            ElseIf Modo = 2 And AntCodArtic <> "" Then
                                If txtCod.Text <> AntCodArtic Then txtNom.Text = vArtic.Nombre
                            End If
                            BloquearTxt txtNom, False
'                            PonerFoco txtNom
                        End If

                        Select Case tipoMov
                            Case "OFE", "PEV", "ALV", "ALR", "FAV", "FTI": If vArtic.TextoVentas <> "" Then vArtic.MostrarTextoVen
                            Case "PEC", "ALC", "FAC": If vArtic.TextoCompras <> "" Then vArtic.MostrarTextoCom
                        End Select
                        txtCod.Text = UCase(txtCod.Text)
                        
                        'devolver si el articulo lleva control de lotes
                        sConLotes = vArtic.TieneNumLote
                        
                        'Si me ha indicado el text donde va el codprove, entonces le pongo
                        If vEmpresa.TieneAnalitica Then
                            If vParamAplic.ModoAnalitica = 0 Then 'ccoste trabajador
                            
                            ElseIf vParamAplic.ModoAnalitica = 1 Then 'ccoste familia
                                'centro de coste
                                txtProv = DevuelveDesdeBDNew(conAri, "sfamia", "codccost", "codfamia", vArtic.Familia, "N")
                                
                            Else
                                txtProv = ""
                            End If
                           
                        Else
                            txtProv = vArtic.Codprove
                        End If
                    End If
                End If
            Else
                txtNom.Text = vArtic.Nombre
            End If
        End If
    End If
    
    Set vArtic = Nothing
End Function


'Lineas de ofertas, pedido y albaranes
'Para decir que hace el F2
Public Sub LabelAyudatxtAux(Indice As Integer, ByRef Lbl As Label)
    Select Case Indice
    Case 3
        'Ver referencia
        Lbl.Caption = "F2  Ver articulo"
    Case 4
        'Consultar precios del articulo
        Lbl.Caption = "F2  Ver precios"
    Case 6, 7
        'Consultar dtos
        Lbl.Caption = "F2  Ver descuentos"
        
    Case Else
        Lbl.Caption = ""
    End Select
End Sub


Public Sub AbrirConsultaPrecio2(Cliente As String, Articulo As String, Fecha As String, Referencia As String)
    'Como desde el formulario
    If IsFormLoaded(frmFacConsultaPrecios2) Then
        MsgBox "No se puede consultar los precios. Esta en la ventana de consulta de precios!!!", vbExclamation
    Else
    
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            
            frmFacPreciosArticuloCliente.Datos = Cliente & "|" & Articulo & "|" & Referencia & "|"
            frmFacPreciosArticuloCliente.Show vbModal
        Else
    
        
            frmFacConsultaPrecios2.Fecha = Fecha
            frmFacConsultaPrecios2.ConsultaDesdeFrm = Cliente & "|" & Articulo & "|"
            frmFacConsultaPrecios2.Show vbModal
        End If
    End If
End Sub



Public Sub AbrirFormularioDtos(Articulo As String)
    frmAlmVerDtos.mCodArtic = Articulo
    frmAlmVerDtos.Show vbModal
End Sub


' ---- [02/11/2009] [LAURA] : al pulsar F2 para abrir articulos en la solapa Documentos|Pedidos
'Puesto en el modulo en Abril 2010
Public Sub AbrirForm_Articulos(Articulo As String)
Dim FrmArt2 As frmAlmArticulosGr
'Dim FrmArt As frmAlmArticulos

    If Trim(Articulo) = "" Then Exit Sub
    If vUsu.Nivel2 = 2 Then Exit Sub  'No tiene permiso
    'Set FrmArt = New frmAlmArticulos
    'FrmArt.DeConsulta = True
    'FrmArt.DatosADevolverBusqueda = "::" & Trim(Articulo)
    'FrmArt.parNumTAb = 6
    'FrmArt.Show vbModal
    'Set FrmArt = Nothing
    
    Set FrmArt2 = New frmAlmArticulosGr
    FrmArt2.DeConsulta = True
    FrmArt2.DatosADevolverBusqueda = "::" & Trim(Articulo)
    FrmArt2.Show vbModal
    Set FrmArt = Nothing
    
End Sub
' -----


'********************************************************************
'********************************************************************
'********************************************************************
'
' COMBOS para el CRM. Asi no lo coje de una tabla
'
'********************************************************************
Public Sub CargaComboMediosCRM(ByRef Co As ComboBox)
    Co.Clear

    Co.AddItem "Teléfono"
    Co.ItemData(Co.NewIndex) = 0
    Co.AddItem "eMail"
    Co.ItemData(Co.NewIndex) = 1
    Co.AddItem "Fax"
    Co.ItemData(Co.NewIndex) = 2
    Co.AddItem "Carta"
    Co.ItemData(Co.NewIndex) = 3
    Co.AddItem "Otros"
    Co.ItemData(Co.NewIndex) = 4
    Co.AddItem "Visita"
    Co.ItemData(Co.NewIndex) = 5
    
End Sub

Public Sub CargaComboEstadoCRM(ByRef Co As ComboBox)
    Co.Clear
    
    Co.AddItem "Pendiente"
    Co.ItemData(Co.NewIndex) = 0
    Co.AddItem "En curso"
    Co.ItemData(Co.NewIndex) = 1
    Co.AddItem "Finalizada"
    Co.ItemData(Co.NewIndex) = 2
    
End Sub




Public Sub GrabarLogDtoSuperior(codtipom As String, Documento As String, Fecha As String, linea As String, Nuevo As Boolean)
Dim C As String

    C = codtipom & Documento & "   Fecha: " & Fecha & "  L:" & linea
    If Nuevo Then
        C = C & "  Nuevo"
    Else
        C = C & "  Modificar"
    End If
    Set LOG = New cLOG
    LOG.Insertar 12, vUsu, C
    Set LOG = Nothing
    
End Sub



Function IsFormLoaded(FormToCheck As Form) As Integer
Dim Y As Integer

For Y = 0 To Forms.Count - 1
If Forms(Y) Is FormToCheck Then
    IsFormLoaded = True
    Exit Function
    End If
Next
IsFormLoaded = False
End Function



'Cuando se sale sin de los frm haciendo alt+a NO hace el lostfocus.
'Para ello forazremos a que recalcule el importe
Public Function RecalculoImporteLineas(ByRef TCan As TextBox, ByRef Timp As TextBox, ByRef TDto1 As TextBox, ByRef TDto2 As TextBox, TipoDto As Byte)

    If Not PonerFormatoDecimal(TCan, 1) Then TCan.Text = ""
    If Not PonerFormatoDecimal(Timp, 2) Then Timp.Text = ""
    If Not PonerFormatoDecimal(TDto1, 4) Then TDto1.Text = ""
    If Not PonerFormatoDecimal(TDto2, 4) Then TDto2.Text = ""
    
    RecalculoImporteLineas = CalcularImporte(TCan.Text, Timp.Text, TDto1.Text, TDto2.Text, TipoDto)

End Function




