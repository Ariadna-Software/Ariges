Attribute VB_Name = "ModFunciones"
Option Explicit

Public Const ValorNulo = "Null"


Public NombreCheck As String

Public Function CompForm(ByRef Formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In Formulario.Controls
        'TEXT BOX
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is MSComm Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 2 And Control.Name = "Text3") Or (Opcion = 3 And Control.Name = "txtAux") Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function



Public Sub limpiar(ByRef Formulario As Form)
    Dim Control As Object

    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub

'-----------------------------------
Public Function ValorParaSQL(Valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim D As Single
Dim i As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else

                    V = CSng(Valor)
                    Valor = V
                End If
            Else

            End If
            Dev = TransformaComasPuntos(CStr(Valor))

        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "H"
            Dev = "'" & Format(Valor, FormatoFecha & " hh:mm:ss") & "'"
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then
            Dev = ValorNulo
        Else
            'Modifica Laura: 04/10/05
            If vtag.TipoDato = "N" Then
                Dev = "0"
            Else
                Dev = "''"
            End If
        End If
    End If
    ValorParaSQL = Dev
End Function



Public Function InsertarDesdeForm(ByRef Formulario As Form, Optional Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or Opcion = 0 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function



Public Function CadenaInsertarDesdeForm(ByRef Formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la función.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = Cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function




' Igual que modifica desde form, pero devuelve el SQL
Public Function CadenaModificaDesdeFormulario(ByRef Formulario As Form, Opcion As Byte) As String
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario2
    CadenaModificaDesdeFormulario = ""
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                                 cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWhere
    
    CadenaModificaDesdeFormulario = Aux
    Exit Function
    
EModificaDesdeFormulario2:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function













Public Function PonerCamposForma(ByRef Formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer


    On Error GoTo EPonerCamposForma


    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In Formulario.Controls
        'TEXTO
        If TypeOf Control Is CommonDialog Then
        
        ElseIf (TypeOf Control Is TextBox) And (Control.visible = True) And (Control.Name = "Text1") Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    
                    If mTag.columna <> "" Then
                        'Debug.Print mTag.columna
                        'If mTag.columna = "porciva3re" Then Stop
                        
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            If mTag.TipoDato = "N" Then
                                If Val(Valor) = 0 Then
                                    Control.Text = ""
                                Else
                                    Control.Text = Valor
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                If IsNull(Valor) Then Valor = 0
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = DBLet(vData.Recordset.Fields(campo))
                    i = 0
                    For i = 0 To Control.ListCount - 1
                        If Control.ItemData(i) = Val(Valor) Then
                            Control.ListIndex = i
                            Exit For
                        End If
                    Next i
                    If i = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    Cad = Err.Description
    Cad = "Poner campos formulario. " & vbCrLf & campo & vbCrLf & Cad & vbCrLf
    MsgBox Cad, vbExclamation
End Function



Public Function PonerCamposFormaFrame(ByRef Formulario As Form, NomTxtBox As String, ByRef vData As Adodc, Optional NomCheck As String, Optional NomCombo As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer

    Set mTag = New CTag
    PonerCamposFormaFrame = False


        For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox And Control.visible = True And Control.Name = NomTxtBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.columna <> "" Then
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True And Control.Name = NomCheck Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox And Control.visible = True And Control.Name = NomCombo Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    i = 0
                    For i = 0 To Control.ListCount - 1
                        If Control.ItemData(i) = Val(Valor) Then
                            Control.ListIndex = i
                            Exit For
                        End If
                    Next i
                    If i = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If

    Next Control

    'Veremos que tal
    PonerCamposFormaFrame = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function


Private Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim RS As Recordset
    ObtenerMaximoMinimo = ""
    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.EOF) Then
            ObtenerMaximoMinimo = CStr(RS.Fields(0))
        End If
    End If
    RS.Close
    Set RS = Nothing
End Function

'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener búsqueda. "
'End Function

'Añado Optional CHECK As String. Para poder realizar las busquedas con los checks
Public Function ObtenerBusqueda(ByRef Formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim Cad As String
Dim SQL As String
Dim tabla As String, columna As String
Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In Formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        If Not paraRPT Then
                            Cad = " MAX(" & mTag.columna & ")"
                        Else
                            Cad = " MAX({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    Else
                        If Not paraRPT Then
                            Cad = " MIN(" & mTag.columna & ")"
                        Else
                            Cad = " MIN({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    End If
                    If Not paraRPT Then
                        SQL = "Select " & Cad & " from " & mTag.tabla
                    Else
                        SQL = "Select " & Cad & " from {" & mTag.tabla & "}"
                    End If
                    SQL = ObtenerMaximoMinimo(SQL)
                    
                    Select Case mTag.TipoDato
                    Case "N"
                        If Not paraRPT Then
                            SQL = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                        Else
                            SQL = "{" & mTag.tabla & "." & mTag.columna & "} = " & TransformaComasPuntos(SQL)
                        End If
                    Case "F"
                        If Not paraRPT Then
                            SQL = mTag.tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Else
                            SQL = "{" & mTag.tabla & "." & mTag.columna & "} = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        End If
                    Case Else
                        If Not paraRPT Then
                            SQL = mTag.tabla & "." & mTag.columna & " = '" & SQL & "'"
                        Else
                            SQL = "{" & mTag.tabla & "." & mTag.columna & "} = '" & SQL & "'"
                        End If
                    End Select
                    SQL = "(" & SQL & ")"
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In Formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Not paraRPT Then
                        SQL = mTag.tabla & "." & mTag.columna & " is NULL"
                    Else
                        SQL = "{" & mTag.tabla & "." & mTag.columna & "} is NULL"
                    End If
                    SQL = "(" & SQL & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In Formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If mTag.Cargado Then
                    Aux = Trim(Control.Text)
                    Aux = QuitarCaracterEnter(Aux) 'Si es multilinea quitar ENTER
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            If Not paraRPT Then
                                tabla = mTag.tabla & "."
                            Else
                                tabla = "{" & mTag.tabla & "."
                            End If
                        Else
                            tabla = ""
                        End If
                        If Not paraRPT Then
                            columna = mTag.columna
                        Else
                            columna = mTag.columna & "}"
                        End If
                    Rc = SeparaCampoBusqueda(mTag.TipoDato, tabla & columna, Aux, Cad, paraRPT)
                    If Rc = 0 Then
                        If SQL <> "" Then SQL = SQL & " AND "
                        If Not paraRPT Then
                            SQL = SQL & "(" & Cad & ")"
                        Else
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If

        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    If mTag.TipoDato <> "T" Then
                        Cad = Control.ItemData(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.tabla & "." & mTag.columna & " = " & Cad
                        Else
                            Cad = "{" & mTag.tabla & "." & mTag.columna & "} = " & Cad
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    Else
                        Cad = Control.List(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.tabla & "." & mTag.columna & " = '" & Cad & "'"
                        Else
                            Cad = "{" & mTag.tabla & "." & mTag.columna & "} = '" & Cad & "'"
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If


        'CHECK
                'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If Control.Value = 1 Then
                        Aux = "1"
                    Else
                        If CHECK <> "" Then
                            CheckBusqueda Control
                            tabla = NombreCheck & "|"
                            If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                        End If
                    End If
                    If Aux <> "" Then
                        If Not paraRPT Then
                            Cad = mTag.tabla & "." & mTag.columna
                        Else
                            Cad = "{" & mTag.tabla & "." & mTag.columna & "} "
                        End If
                        
                        Cad = Cad & " = " & Aux
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function


Public Function ModificaDesdeFormulario(ByRef Formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                                 cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWhere
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Sub FormateaCampo(vTex As TextBox)
'devuelve el valor del control vText.text formateado: 12 -> "0012"
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub

Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim Cad As String
On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    i = 0
    cont = 1
    Cad = ""
    Do
        J = i + 1
        i = InStr(J, Cadena, "|")
        If i > 0 Then
            If cont = Orden Then
                Cad = Mid(Cadena, J, i - J)
                i = Len(Cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until i = 0
    RecuperaValor = Cad
End Function




'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef Formulario As Form)
Dim i As Integer
Dim J As Integer

On Error GoTo EPonerOpcionesMenuGeneral

'Añadir, modificar y borrar deshabilitados si no nivel
With Formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For i = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(i).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(i).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(i).Enabled = False
            End If
        End If
    Next i

    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next

    On Error Resume Next

    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnNuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnNuevo.Enabled = False

    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False

    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    
    J = Val(.mnLineas.HelpContextID)
    If J < vUsu.Nivel Then .mnLineas.Enabled = False
    
    On Error GoTo 0
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



Public Function BLOQUEADesdeFormulario(ByRef Formulario As Form, Optional Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or Opcion <> 1 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                             cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistro(cadTabla As String, cadWhere As String) As Boolean
Dim Aux As String
On Error GoTo EBloqueaRegistro

        BloqueaRegistro = False
        
        Aux = "SELECT * FROM " & cadTabla
        Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BloqueaRegistro = True
        
EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef Formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Name = "Text1" Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "Insert into zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & ComprobarComillas(AuxDef) & """)"
        conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Private Function ComprobarComillas(Cad As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, Cad, """")
        If i > 0 Then
            Aux = Mid(Cad, 1, i - 1) & "\"
            Cad = Aux & Mid(Cad, i)
            J = i + 2
        End If
    Loop Until i = 0
    ComprobarComillas = Cad
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
        SQL = "DELETE from zbloqueos where codusu=" & vUsu.codigo & " and tabla='" & mTag.tabla & "'"
        conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function



Public Function BloqueoManual(cadTabla As String, cadWhere As String, Optional OcultarMsg As Boolean) As Boolean
Dim Aux As String

On Error GoTo EBLOQ
    BloqueoManual = False
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.codigo & ",'" & cadTabla
        Aux = Aux & "',""" & cadWhere & """)"
        conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            If Not OcultarMsg Then MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTabla As String) As Boolean
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next

        SQL = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.codigo & " and tabla='" & cadTabla & "'"
        conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function


'====================== LAURA

Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function


Public Function QuitarCero(Valor As String) As String
    On Error Resume Next
    
    If Valor <> "" Then
        If CSng(Valor) = 0 Then
            QuitarCero = ""
        Else
            QuitarCero = Valor
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function CalcularImporte(cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CCur(cantidad) * CCur(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    '// Enero 2009.  Hacia mal el redondeo pq ahora cantidad lleva decimales
    '   Ponemos Round2
    vImp = Round2(vImp, 2)
    CalcularImporte = CStr(vImp)
End Function

'Redondeo a 4 digitos
Public Function CalcularImporte4(cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CCur(cantidad) * CCur(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    vImp = Round(vImp, 4)
    CalcularImporte4 = CStr(vImp)
End Function



Public Function CalcularImporteSng(cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran,
'donde PRECIO es sng                                          *********************** MAYO 2009
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Single
Dim vDto1 As Single, vDto2 As Single
Dim vPre As Single
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CSng(cantidad) * CSng(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    '// Enero 2009.  Hacia mal el redondeo pq ahora cantidad lleva decimales
    '   Ponemos Round2
    vImp = Round2(vImp, 2)
    CalcularImporteSng = CStr(vImp)
End Function





Public Function CalcularDto(Importe As String, Dto As String) As String
'devuelve el Dto% del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vDto As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    Dto = ComprobarCero(Dto)
    
    vImp = CCur(Importe)
    vDto = CCur(Dto)
    
    vImp = ((vImp * vDto) / 100)
    'vImp = Round(vImp, 2)
    
    CalcularDto = CStr(vImp)
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularNumBultos2(cantidad As Currency, UdsCaja As Integer) As Long
Dim numUds As Long 'unidades
    
    If UdsCaja > 0 Then
        '- calcular los bultos q necesitamos para la cantidad
        numUds = Int(cantidad / UdsCaja)
        If cantidad Mod UdsCaja > 0 Then
            numUds = numUds + 1
        ElseIf cantidad > Val(UdsCaja * numUds) Then
             numUds = numUds + 1
        End If
        
        
        If numUds = 0 And cantidad <> 0 Then numUds = numUds + 1
    End If
    
    CalcularNumBultos2 = numUds
End Function


'Si pone algo en DevuelveImporte en lugar del msg metera en esa cadena el importe
Public Sub ComprobarCobrosCliente(codClien As String, FechaDoc As String, Optional DevuelveImporte As String)
'Comprueba en la tabla de Cobros Pendientes (scobro) de la Base de datos de Contabilidad
'si el cliente tiene alguna factura pendiente de cobro que ha vendido
'con fecha de vencimiento anterior a la fecha del documento: Oferta, Pedido, ALbaran,...
Dim SQL As String, vWhere As String
Dim Codmacta As String
Dim RS As ADODB.Recordset
Dim cadMen As String
Dim ImporteCred As Currency
Dim Importe As Currency
Dim Impaux As Currency

    Set RS = New ADODB.Recordset
    ImporteCred = 0
    'Obtener la cuenta del cliente de la tabla sclien en Ariges
    SQL = "Select nomclien,codmacta,limcredi,clivario from sclien where codclien=" & codClien
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        SQL = ""
    Else
        'CodClien = CodClien & " - " & sql
        If DBLet(RS!Clivario, "N") = 1 Then
            SQL = ""
        Else
            codClien = codClien & " - " & RS!Nomclien
            ImporteCred = DBLet(RS!limcredi, "N")
            If ImporteCred > 0 Then codClien = codClien & "   Límite credito: " & Format(ImporteCred, FormatoImporte)
            Codmacta = RS!Codmacta
        End If
    End If
    RS.Close
    If SQL = "" Then Exit Sub
    
    'AHORA FEBRERO 2010
    If vParamAplic.ContabilidadNueva Then
    
        SQL = "SELECT cobros.* FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
        vWhere = " WHERE cobros.codmacta = '" & Codmacta & "'"
        vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
        'Antes mayo 2010
        'vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
        vWhere = vWhere & " AND recedocu=0 ORDER BY fecfactu, numfactu"
        SQL = SQL & vWhere
    
    
    Else
    
        SQL = "SELECT scobro.* FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        vWhere = " WHERE scobro.codmacta = '" & Codmacta & "'"
        vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
        'Antes mayo 2010
        'vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
        vWhere = vWhere & " AND recedocu=0 ORDER BY fecfaccl, codfaccl"
        SQL = SQL & vWhere
    End If
    'Lee de la Base de Datos de CONTABILIDAD
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    While Not RS.EOF
    
        
        If Val(RS!recedocu) = 1 Then
            Impaux = 0
        Else
            'NO esta recibido. Si tiene diferencia
            Impaux = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
    
        End If
    '    End If
        If Impaux <> 0 Then Importe = Importe + Impaux
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        If Importe > 0 Then
        
            If DevuelveImporte <> "" Then
                'Meto aqui el importer
                DevuelveImporte = CStr(Importe)
            Else
                cadMen = "El Cliente tiene facturas vencidas con valor de: " & Format(Importe, FormatoImporte) & " €."
                If ImporteCred > 0 Then cadMen = cadMen & vbCrLf & "Límite crédito: " & Format(ImporteCred, FormatoImporte) & " €."
                cadMen = cadMen & vbCrLf & "¿Desea Ver Detalle?"
                If MsgBox(cadMen, vbYesNo + vbQuestion + vbDefaultButton2, "Cobros Pendientes") = vbYes Then
                    'Mostrar los detalles de los cobros pendientes
                    frmMensajes.cadWhere = vWhere
                    frmMensajes.vCampos = codClien
                    frmMensajes.OpcionMensaje = 1
                    frmMensajes.Show vbModal
                End If
            End If
        End If
    
    
End Sub

'Tipoiva:  0 normal  1 RE    2,3 Extento e intracom
Public Sub RiesgoCliente(Cliente As Long, TipoIVA As Byte, Fecha As Date, ByRef ImporteTesoreria As Currency, ByRef ImporteAlbaranes As Currency, ByRef RIVA As ADODB.Recordset)
Dim RA As ADODB.Recordset
Dim SQL As String
Dim Aux As Currency
Dim VieneCargadoIVA As Boolean
Dim codigo As Integer

    Set RA = New ADODB.Recordset
    
    VieneCargadoIVA = True
    If RIVA Is Nothing Then
        VieneCargadoIVA = False
        'Cargo el IVA
        Set RIVA = New ADODB.Recordset
        
        RIVA.Open "Select * from tiposiva ", ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
        
            
    End If

   
    'marzo 2011
    'El riesgo es el select sum del importe + gasto paraa la cuenta contable
    SQL = "Select codmacta from sclien WHERE codclien=" & CStr(Cliente)
    RA.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "NO"
    If Not RA.EOF Then
        If Not IsNull(RA!Codmacta) Then SQL = RA!Codmacta
    End If
    RA.Close
    
    
    SQL = " where codmacta = " & DBSet(SQL, "T")
    SQL = "Select sum(impvenci),sum(if(gastos is null,0,gastos)) from #####  " & SQL
   

    
    If vParamAplic.ContabilidadNueva Then
        SQL = Replace(SQL, "#####", "cobros")
    Else
        SQL = Replace(SQL, "#####", "scobro")
    End If
    
    RA.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = 0
    If Not RA.EOF Then Aux = DBLet(RA.Fields(0), "N") + DBLet(RA.Fields(1), "N")
    RA.Close
    ImporteTesoreria = Aux
    'Antes
    'SQL = "i"
    'ComprobarCobrosCliente CStr(Cliente), CStr(Fecha), SQL
    'If SQL = "i" Then SQL = "0"
    'ImporteTesoreria = CCur(SQL)
   
   
 
    SQL = "select codigiva,sum(importel) from scaalb,slialb,sartic where scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar "
    SQL = SQL & "and slialb.codartic=sartic.codartic and scaalb.codtipom<>'ART' and factursn=1 and codclien=" & CStr(Cliente) & " group by 1"
    RA.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ImporteAlbaranes = 0
    While Not RA.EOF
        codigo = RA!codigiva
        If TipoIVA = 1 Then
            'Recargo equivalencia
            codigo = vParamAplic.DevuleveTipoIVA_RE(codigo)
        ElseIf TipoIVA > 1 Then
            'Sera 0
            codigo = -1
        Else
            'nada
        End If
        RIVA.Find "codigiva = " & codigo, , adSearchForward, 1
        If RIVA.EOF Then
            Aux = 0
        Else
            Aux = RIVA!PorceIVA
        End If
        If IsNull(RA.Fields(1)) Then
            Aux = 0
        Else
            Aux = (Aux / 100) + 1
        End If
        Aux = Round2(Aux * RA.Fields(1), 2)
        ImporteAlbaranes = ImporteAlbaranes + Aux
        RA.MoveNext
    Wend
    RA.Close
    Set RA = Nothing
    
    If Not VieneCargadoIVA Then
        RIVA.Close
        Set RIVA = Nothing
    End If
End Sub






'--------------------------------------------------------

Public Function EsArticuloVarios(codArtic As String) As Boolean
Dim devuelve As String

    EsArticuloVarios = False
    devuelve = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", codArtic, "T")
    
    If devuelve = "1" Or devuelve = "2" Then 'Es Articulo de Varios y podemos modificar la Denominación del Articulo
        EsArticuloVarios = True
    Else
        EsArticuloVarios = False
    End If
End Function


Public Function EsClienteVarios(vCodClien As String) As Boolean
'Devuelve true si es un cliente de varios
Dim devuelve As String

    EsClienteVarios = False
    devuelve = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", vCodClien, "N")
    If devuelve <> "" Then EsClienteVarios = CBool(devuelve)
    'Es cliente de varios Y podemos recuperar de sclvar los datos
    'del cliente por el NIF
End Function



Public Function EsClienteBloqueado(codClien As String, DesdeOFertasPedidos As Boolean, OcultarMsg As Boolean) As Boolean
'NUEVO
'Diciembre 2010
'Si esta en ofertas / pedidos CLIENTE buscaremos en la tabla el campo: clioferped
' Si es 1 tb bloquea la creacion de ofertas / pedidos


'devuelve true si el cliente esta bloqueado
'si la situación del cliente es distinta de NORMAL(codsitua=0) entonces
'mostrar un mensaje con la situación especial del cliente
Dim Tipo As String
Dim devuelve As String

    On Error GoTo EBloqueado
    EsClienteBloqueado = False
    
    
    If DesdeOFertasPedidos Then
        Tipo = "clioferped"
    Else
        Tipo = "tipositu"
    End If
    
    devuelve = "select ssitua.* from  where "
    devuelve = DevuelveDesdeBDNew(conAri, "sclien,ssitua", "nomsitua", "sclien.codsitua=ssitua.codsitua AND codclien ", codClien, "N", Tipo)
    If Tipo = 0 Then
        'NO BLOQUEA
        
    Else
        If Not OcultarMsg Then MsgBox UCase("Cliente Bloqueado por: ") & vbCrLf & devuelve, vbInformation, "Situación Especial del Cliente."
        EsClienteBloqueado = True
    
    End If
    
EBloqueado:
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function EsProveedorVarios(Codprove As String) As Boolean
Dim devuelve As String

    EsProveedorVarios = False
    devuelve = DevuelveDesdeBD(conAri, "provario", "sprove", "codprove", Codprove, "N")
    If devuelve <> "" Then EsProveedorVarios = CBool(devuelve)
    'Es proveedor de varios Y podemos recuperar de ????
End Function


Public Function ObtenerNSerieSiguiente(cadNSerie As String) As String
'IN -> cadNSerie: cadena con el Nº Serie de Tipo: "0000-12-0011"
'OUT -> RETURN: cadena con el sig. NºSerie : "0000-12-0012"
Dim NumAux As String, numAnt As String
Dim NumAux2 As String
Dim i As Integer

    On Error Resume Next
    
    NumAux = cadNSerie
    numAnt = ""
    'Quitar los cararacter '-' y quedarse con la parte dcha
    i = InStr(1, NumAux, "-")
    While Not i = 0
        numAnt = numAnt & Mid(NumAux, 1, i)
        NumAux = Mid(NumAux, i + 1, Len(NumAux) - i)
        i = InStr(1, NumAux, "-")
    Wend
    
    If NumAux <> "" Then 'Hay q coger la parte derecha del - : 0011
        i = Len(NumAux)
        If IsNumeric(NumAux) Then
            NumAux = CStr(NumAux + 1)
            While Len(NumAux) < i
                NumAux = "0" & NumAux
            Wend
        Else
        'Coger el nº mas a la derecha, incrementarlo y concatenarlo con el principio
            NumAux2 = Mid(NumAux, i, Len(NumAux))
            While IsNumeric(NumAux2)
                i = i - 1
                NumAux2 = Mid(NumAux, i, Len(NumAux))
            Wend
            NumAux2 = Right(NumAux2, Len(NumAux2) - 1)
            numAnt = numAnt & Mid(NumAux, 1, i)
            NumAux = CStr(NumAux2 + 1)
            While Len(NumAux) < Len(NumAux2)
                NumAux = "0" & NumAux
            Wend
        End If
        
        If numAnt <> "" Then
            ObtenerNSerieSiguiente = numAnt & NumAux
        Else
            ObtenerNSerieSiguiente = NumAux
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PonerTrabajadorConectado(NomTraba As String) As String
'Pone en el campo del Form "Realizada Por" el trabajador que esta conectado en ese momento
'OUT: codTraba, NomTraba
Dim devuelve As String

    On Error Resume Next

    NomTraba = "nomtraba"
    devuelve = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "login", vUsu.Login, "T", NomTraba)
    If devuelve <> "" Then
        PonerTrabajadorConectado = Format(devuelve, "0000") 'Cod. Trabajador
    Else
        PonerTrabajadorConectado = ""
        NomTraba = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function PonerAlmacen(codAlm As String) As String
'Comprueba si existe el Almacen y lo pone en el Text
Dim devuelve As String
    
    On Error Resume Next

    If codAlm = "" Then
        MsgBox "Debe introducir el Almacen.", vbInformation
    Else
        devuelve = DevuelveDesdeBDNew(conAri, "salmpr", "codalmac", "codalmac", codAlm, "N")
        If devuelve = "" Then
            MsgBox "No existe el Almacen: " & Format(codAlm, "000"), vbInformation
            PonerAlmacen = ""
        Else
            PonerAlmacen = Format(codAlm, "000")
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


'=============================================================================
'==================== REPARACIONES ===========================================

Public Sub ComprobarReparaciones(Modo As Byte, numSerie As String, codArtic As String)
Dim numRep As Integer

    'Comprobar si ya esta en Reparacion
    If Modo = 3 Then ComprobarSiReparandose numSerie, codArtic
    'Comprobar cuantas veces se ha reparado ya el articulo(ver historico Reparaciones)
    numRep = ComprobarNumRepHco(numSerie, codArtic)
    If numRep > 0 Then
        MsgBox "Este aparato ya ha sido reparado " & numRep & " veces.", vbInformation
    End If
End Sub



Public Function ComprobarSiReparandose(numSerie As String, codArtic As String) As Boolean
'Comprueba si ya el Articulo se esta reparando, es decir si existe un registro
' en la tabla scarep
'IN -> numSerie, codArtic
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(conAri, "scarep", "numrepar", "numserie", numSerie, "T", , "codartic", codArtic, "T")
    If devuelve <> "" Then
        MsgBox "Este aparato ya esta en Reparación.", vbInformation
        ComprobarSiReparandose = True
    Else
        ComprobarSiReparandose = False
    End If
End Function


Public Function ComprobarNumRepHco(numSerie As String, codArtic As String) As Integer
'Comprueba cuantas veces se ha reparado ya el articulo
'Ver cuantos registros existen en la tabla de historico Reparaciones (schrep)
'IN -> numserie, codartic
'RETURN -> Nº Reparaciones
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ENumRep

    SQL = " SELECT count(numrepar) FROM schrep "
    SQL = SQL & " WHERE numserie=" & DBSet(numSerie, "T") & " and "
    SQL = SQL & " codartic=" & DBSet(codArtic, "T")

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ComprobarNumRepHco = RS.Fields(0).Value
    Else
        ComprobarNumRepHco = 0
    End If
    
    RS.Close
    Set RS = Nothing
    
ENumRep:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Public Function ObtenerLetraSerie(tipMov As String) As String
'Devuelve la letra de serie asociada al tipo de movimiento
Dim LEtra As String

    On Error Resume Next
    
    LEtra = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", tipMov, "T")
    If LEtra = "" Then MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
    ObtenerLetraSerie = LEtra
End Function


Public Function ObtenerPoblacion(CPostal As String, ByRef provin As String) As String
'IN: "cpostal"
'OUT: en "provin" devolvemos la provincia
'     en ObtenerPoblacion devolvemos la poblacion
Dim devuelve As String

    On Error GoTo EPoblacion

    If CPostal <> "" Then
        devuelve = DevuelveDesdeBDNew(conAri, "scpostal", "provincia", "cpostal", CPostal, "T")
        ObtenerPoblacion = devuelve 'Nombre Poblacion
        If devuelve <> "" Then 'Nombre Provincia
            provin = DevuelveDesdeBDNew(conAri, "scpostal", "provincia", "cpostal", Mid(CPostal, 1, 2), "T")
        Else
            provin = ""
            MsgBox "No existe el CPostal " & CPostal, vbInformation
        End If
    Else
        ObtenerPoblacion = ""
        provin = ""
    End If
    
EPoblacion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener Población", Err.Description
End Function


Public Sub ObtenerCtasBancoPropio2(banPr As String, ctaBan As String, ctaCble As String)
'obtener la cuenta bancaria y la cuenta contable del banco propio
'(IN) banPr: cod. banco propio
'(OUT) ctaBan: cuenta bancaria
'(OUT) ctaCble: cuenta contable
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Aux As String

    ctaBan = ""
    ctaCble = ""

    SQL = "SELECT codbanco,codsucur,digcontr,cuentaba,codmacta"
    SQL = SQL & " from sbanpr where codbanpr=" & banPr

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        Aux = Right("0000" & DBLet(RS!codbanco, "T"), 4)
        ctaBan = Aux & "-"
        Aux = Right("0000" & DBLet(RS!codsucur, "T"), 4) & "-"
        ctaBan = ctaBan & Aux
        ctaBan = ctaBan & DBLet(RS!digcontr, "T") & "-" & DBLet(RS!cuentaba, "T")
        ctaCble = DBLet(RS!Codmacta, "T")
        'obtener el nombre de la cuenta contable
        SQL = ""
        SQL = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", ctaCble, "T")
        If SQL <> "" Then ctaCble = ctaCble & "-" & SQL
    End If
    Set RS = Nothing
End Sub



Public Function ObtenerSQLcomponentes(cadWhere As String) As String
'Obtiene la consulta SQL que selecciona los articulos con nº de serie
'agrupados por tipo de articulo
Dim SQL As String

    SQL = "Select distinct sserie.codtipar, nomtipar, count(numserie) as cantidad "
    SQL = SQL & "FROM sserie INNER JOIN stipar ON sserie.codtipar=stipar.codtipar "
    SQL = SQL & cadWhere
    SQL = SQL & " GROUP by codtipar "
    
    ObtenerSQLcomponentes = SQL
End Function



Public Function ComprobarStock(codArtic As String, codAlmac As String, Cant As String, CodTipMov As String) As Boolean
'Comprueba si el Articulo existe en el Almacen Origen y si hay
'stock suficiente para poder realizar el traspaso
Dim vStock As String
Dim vArtic As CArticulo
Dim B As Boolean

    Set vArtic = New CArticulo
    B = vArtic.Existe(codArtic)
    If B Then
        B = vArtic.ExisteEnAlmacen(codAlmac, vStock)
        If B Then
            B = ComprobarHayStock(CSng(vStock), CSng(Cant), codArtic, vArtic.Nombre, CodTipMov)
'            If Not ComprobarHayStock(CSng(vStock), CSng(cant), codArtic, vArtic.Nombre, CodTipMov) Then
'                b = False
'            Else
'                b = True
'            End If
        End If
    End If
    Set vArtic = Nothing
    ComprobarStock = B
End Function



Public Function ObtenerPrecioSinIVAvarios(codArtic As String, Precio As String) As Currency
Dim vArtic As CArticulo
Dim PreuSinIVA  As Currency

'    On Error GoTo ErrTotal
'
''    If sPorce <> "" Then curPorce = ImporteFormateado(sPorce)
'    If Precio <> "" Then PreuConIVA = ImporteFormateado(Precio) 'precio con iva

    Set vArtic = New CArticulo
    If vArtic.LeerDatos(codArtic) Then
        'precio con iva del articulo
        PreuSinIVA = vArtic.ObtenerPrecioSinIVA(Precio)
    Else
        PreuSinIVA = CCur(ComprobarCero(Precio))
    End If

'
'
'    curPorce = curPorce / 100
'    curImporte = curImporte / (1 + curPorce) 'importe sin iva
'    curCuota = Round((curPorce * curImporte), 2)
'    curImporte = Round(curImporte, 2)
'
'    'valores que devuelve: Importe sin iva, cuota de iva
'    ImporteSinIVA = Format(curImporte, FormatoImporte)
'    sCuota = Format(curCuota, FormatoImporte)
'
'    Exit Function


'    Set vArtic = New CArticulo
'    If vArtic.LeerDatos(codArtic) Then
'        'precio con iva del articulo
'        PreuIVA = vArtic.ObtenerPrecioConIVA
'    End If
'
'
'    'El precio con IVA calculado a partir del importe del articulo no coincide con el
'    'precio con IVA introducido en la linea.
'    'recalculamos el importe del articulo SIN iva (se modifica precio original del artic)
'    If Round(PreuIVA, 2) <> Round(CCur(Precio), 2) Then
'        If PreuIVA <> 0 Then
'            PreuIVA = Round((vArtic.PrecioVenta * CCur(Precio)) / PreuIVA, 4)
'        Else
'            PreuIVA = Round((CCur(Precio) * 100) / (100 + vArtic.ObtenerPorceIVA), 4)
'        End If
'    Else
'        PreuIVA = vArtic.PrecioVenta
'    End If
    Set vArtic = Nothing
    ObtenerPrecioSinIVAvarios = PreuSinIVA
End Function




 



Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim Cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function CApos(texto As Variant) As String
    Dim i As Integer
    i = InStr(1, texto, "'")
    If i = 0 Then
        CApos = texto
    Else
        CApos = Mid(texto, 1, i - 1) & "\'" & Mid(texto, i + 1, Len(texto) - i)
    End If
    '-- Ya que estamos transformamos las Ñ
    texto = CApos
    i = InStr(1, texto, "¥")
    If i = 0 Then
        CApos = texto
    Else
        CApos = Mid(texto, 1, i - 1) & "Ñ" & Mid(texto, i + 1, Len(texto) - i)
    End If
    '-- Y otra más
    texto = CApos
    i = InStr(1, texto, "¾")
    If i = 0 Then
        CApos = texto
    Else
        CApos = Mid(texto, 1, i - 1) & "Ñ" & Mid(texto, i + 1, Len(texto) - i)
    End If
    '-- Seguimos con transformaciones
    texto = CApos
    i = InStr(1, texto, "¦")
    If i = 0 Then
        CApos = texto
    Else
        CApos = Mid(texto, 1, i - 1) & "ª" & Mid(texto, i + 1, Len(texto) - i)
    End If
End Function


Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim ent As Integer
Dim Cad As String
  
  ' Comprobaciones
  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If
  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If
  
  ' Redondeo.
  Cad = "0"
  If NumDigitsAfterDecimals <> 0 Then Cad = Cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Format(Number, Cad)
  
End Function



Public Function CalcularPorcentaje(Importe As Currency, Porce As Currency, NumDecimales As Long) As Variant
'devuelve el valor del Porcentaje aplicado al Importe
'Ej el 16% de 120 = 19.2
'Dim vImp As Currency
'Dim vDto As Currency
    
    On Error Resume Next
'
'    Importe = ComprobarCero(Importe)
'    Dto = ComprobarCero(Dto)
'
'    vImp = CCur(Importe)
'    vDto = CCur(Dto)
    
    
    'vImp = Round(vImp, 2)
    
    CalcularPorcentaje = Round2((Importe * Porce) / 100, NumDecimales)
    
    If Err.Number <> 0 Then Err.Clear
End Function




Public Function ArticuloTieneMargen(codArt As String) As Boolean
Dim Cad As String

    'Comprobar que el artículo tiene margen comercial
    Cad = DevuelveDesdeBDNew(conAri, "sartic", "margecom", "codartic", codArt, "T")
    If Cad = "" Then
        Cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
        Cad = Cad & "El artículo no tiene margen comercial para calcular nuevos precios."
        MsgBox Cad, vbExclamation
        ArticuloTieneMargen = False
        Exit Function
    End If
    
    
'    'comprobar que las tarifas del articulo tienen margen comercial
'    cad = "SELECT count(*)"
'    cad = cad & " FROM slista INNER JOIN starif ON slista.codlista = starif.codlista "
'    cad = cad & " WHERE slista.codartic=" & DBSet(codArt, "T") & " AND  isnull(margecom)"
'    If RegistrosAListar(cad) > 0 Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El artículo tiene tarifas sin %PVP necesario para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
    
    ArticuloTieneMargen = True
    
End Function


Public Function ActualizacionAutomaticaMargen(codArtic As String)
Dim PrecioVen As Currency
Dim PrecioCom As Currency
Dim RN As ADODB.Recordset
Dim Cad As String

    On Error GoTo eActualizacionAutomaticaMargen

    Cad = "select margecom,preciouc,preciove,nomartic from sartic where codartic=" & DBSet(codArtic, "T")
    Set RN = New ADODB.Recordset
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RN.EOF Then
        'NUUUUNCA deberia pasar
        MsgBox "Error grave, muy grave", vbCritical
    Else
        If Not IsNull(RN!margecom) Then
            If RN!precioUC > 0 Then
                PrecioCom = Round2(RN!PrecioVe / RN!precioUC, 4)
                PrecioCom = (PrecioCom - 1) * 100
                If Val(PrecioCom) > 999 Then
                    MsgBox "margen calculado:" & Format(PrecioCom, FormatoPrecio) & "   superior al 999%. No se actualizara"
                Else
                    If PrecioCom < 0 Then
                        MsgBox "Margen negativo. Pasa a ser 0", vbExclamation
                        PrecioCom = 0
                    End If
                    PrecioCom = Round2(PrecioCom, 2)
                    Cad = "UPDATE sartic set margecom=" & DBSet(PrecioCom, "N") & " WHERE codartic = " & DBSet(codArtic, "T")
                    conn.Execute Cad
                    Cad = "Cambio en el margen del articulo. " & vbCrLf & vbCrLf
                    Cad = Cad & "Código: " & codArtic & vbCrLf
                    Cad = Cad & "Descrip: " & RN!NomArtic & vbCrLf
                    Cad = Cad & "% Anteor:        " & Format(RN!margecom, FormatoImporte) & vbCrLf & vbCrLf
                    Cad = Cad & "****  ACTUAL: " & Format(PrecioCom, FormatoImporte) & "     ****"
                    MsgBox Cad, vbInformation
                End If
            End If
        End If
    End If
    RN.Close
    
eActualizacionAutomaticaMargen:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RN = Nothing
End Function



Public Function TotalRegistros(vSQL As String, Optional vBD As Byte) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    If vBD = conConta Then 'Accede a BD de contabilidad
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    TotalRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TotalRegistros = RS.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function

'---------------------------------------------------------------------------------
'
'       Para buscar en los checks con las dos opciones de true y false
'
'A partir de un check cualquiera devolvera nombre e indice, si tiene. Si no sera ()
Public Sub CheckBusqueda(ByRef CH As CheckBox)
    NombreCheck = ""
    NombreCheck = CH.Name & "("
    On Error Resume Next
    NombreCheck = NombreCheck & CH.Index
    If Err.Number <> 0 Then Err.Clear
    NombreCheck = NombreCheck & ")"
End Sub



Public Sub CheckCadenaBusqueda(ByRef CH As CheckBox, ByRef CadenaCHECKs As String)
        CheckBusqueda CH
        If InStr(1, CadenaCHECKs, NombreCheck) = 0 Then CadenaCHECKs = CadenaCHECKs & NombreCheck & "|"
End Sub




'---------------------------------------------------------------------------------
'
'       Las tabla reparaciones esta relacionada, sin FOREING KEY con
'       SAT, tipoave,trabajorealizado
'       Para saber si se puede eliminar alguno de estos
'       mantenimientos entonces trendrmos esta funcion
'
'       Opcion
'           1:  sat
'           2:  tipoave
'           3:  trabajaorealizado
Public Function SePuedeEliminarRelReparacione(Opcion As Byte, codigo As String) As Boolean
Dim CA As String
Dim C2 As String

    SePuedeEliminarRelReparacione = False
    If Opcion = 1 Then
        'SAT
        CA = "codman"
    Else
        If Opcion = 2 Then
            CA = "codavi" 'Deberia haber sido AVE de averia, no avi
        Else
            CA = "codtrabajo"
        End If
    End If
    'Miramos primero en scarep
    C2 = DevuelveDesdeBDNew(conAri, "scarep", "numrepar", CA, codigo, "N")
    If C2 <> "" Then Exit Function
        
        
    'Ahora miraremos en hco reparaciones
    C2 = DevuelveDesdeBDNew(conAri, "schrep", "numrepar", CA, codigo, "N")
    If C2 <> "" Then Exit Function

    
    SePuedeEliminarRelReparacione = True
End Function

Public Function SugerirCodAutomatico(marca As String, categoria As String, modelo As String, Formato As String) As String
    '-- SugerirCodAtomatico:
    '   Esta función se utiliza en el marco del parámetro descriptores y sirve, al igual que se montaba un descriptor
    '   automático a partir de las descripciones de los campos de marca, categoria, modelo y formato; hacer lo propio
    '   pero con el código. Con el siguiente formato
    '   MMMMCCCCmmffXXXX -> M=marca, C=categoria, m=modelo, f=formato, x=un ordinal para el código
    Dim inferior As String
    Dim superior As String
    Dim comun As String
    Dim codigo As String
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Valor As Integer
    '-- Primero trimeamos los valores por si acaso.
    marca = Left(Trim(marca) & "0000", 4)
    categoria = Left(Trim(categoria) & "0000", 4)
    modelo = Left(Trim(modelo) & "00", 2)
    Formato = Left(Trim(Formato) & "00", 2)
    '--
    comun = marca & categoria & modelo & Formato
'    inferior = comun & "0000"
'    superior = comun & "9999"
'
'    SQL = "select max(codartic) from sartic where" & _
'            " codartic >= '" & inferior & "'" & _
'            " and codartic <= '" & superior & "'"
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly
'    '-- por defecto el código es:
'    codigo = comun & "0001"
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then
'            If Not IsNumeric(Right(RS.Fields(0), 4)) Then
'                MsgBox "La cola de código: " & RS.Fields(0) & " no es numérica. No puedo sugerir el código siguiente", vbExclamation
'                codigo = ""
'            Else
'                Valor = Val(Right(RS.Fields(0), 4)) + 1
'                codigo = comun & Format(Valor, "0000")
'            End If
'        End If
'    End If
'    SugerirCodAutomatico = codigo
    SugerirCodAutomatico = comun
End Function

Public Function CambiaTagDescriptores(ByRef txt As TextBox, descriptor As String) As String
    '-- Cambia el comienzo del tag del descriptor en el tag, para que cuando diga xxx no exista, aparezca
    '   la etiqueta correcta.
    Dim pos As Integer
    Dim ntag As String
    ntag = txt.Tag
    pos = InStr(1, ntag, "|")
    If pos Then
        ntag = descriptor & Mid(ntag, pos, (Len(ntag) - pos) + 1)
    End If
    txt.Tag = ntag
    CambiaTagDescriptores = ntag
End Function


'                                                                       CINCO DECIMALES
Public Function ArticuloConTasaReciclado(ArticuloLinea As String, ByRef ImporteSng As Single) As Boolean
Dim RT As ADODB.Recordset
Dim SQL As String
        On Error GoTo EArticuloConTasaReciclado
        ArticuloConTasaReciclado = False
        SQL = "select tasareciclado from sunida,sartic where sunida.codunida =sartic.codunida and sartic.codartic=" & DBSet(ArticuloLinea, "T")
        Set RT = New ADODB.Recordset
        RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RT.EOF Then
            If Not IsNull(RT!tasareciclado) Then
                ImporteSng = RT!tasareciclado
                ArticuloConTasaReciclado = True
            End If
        End If
        RT.Close
        Set RT = Nothing
        Exit Function
EArticuloConTasaReciclado:
    MuestraError Err.Number, Err.Description, "Calculando tasa reciclado."
    Set RT = Nothing
End Function



Public Function DevuelveUltimoAlmacen(tabla As String, Where As String) As Integer
Dim C As String
Dim RS As ADODB.Recordset

    DevuelveUltimoAlmacen = -1
    C = "Select codalmac FROM " & tabla & Where & " ORDER BY numlinea DESC"
    Set RS = New ADODB.Recordset
    RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then DevuelveUltimoAlmacen = CInt(RS.Fields(0))
    End If
    RS.Close
    Set RS = Nothing
End Function




'Le paso varuiables por si acaso quiero sacar la function de aqui
'Direccion de envio o para departamentos.
'los departamentos, aunque tienen clave tab podemos comporbarlo
Public Function PuedeEliminarDirecEnvio(DirecionEnvio As Boolean, ByRef Cliente As String, ByRef codigo As Integer) As Boolean
Dim Aux As String
Dim Cad As String
Dim Resul As String

    If DirecionEnvio Then
        Aux = "coddiren"
        Resul = "con esta direccion de envio"
    Else
        Aux = "coddirec"
        Resul = DevuelveTextoDepto(False)
        Resul = "asociados a " & Resul
    End If
    PuedeEliminarDirecEnvio = False
    'Busco en las tablas
    'PEDIDOS
    Cad = DevuelveDesdeBDNew(1, "scaped", "count(*)", "codclien", Cliente, "N", "", Aux, CStr(codigo), "N")
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Existen pedidos " & Resul, vbExclamation
        Exit Function
    End If
    
    Cad = DevuelveDesdeBDNew(1, "scapre", "count(*)", "codclien", Cliente, "N", "", Aux, CStr(codigo), "N")
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Existen ofertas " & Resul, vbExclamation
        Exit Function
    End If
        
    Cad = DevuelveDesdeBDNew(1, "scaalb", "count(*)", "codclien", Cliente, "N", "", Aux, CStr(codigo), "N")
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Existen albaranes " & Resul, vbExclamation
        Exit Function
    End If
    
    
    
    'Las facturas es mas complicado, si es coddiren
    If DirecionEnvio Then
        Cad = " coddiren = " & codigo & " AND (codtipom, numfactu, fecfactu) IN "
        Cad = Cad & "(Select codtipom ,numfactu ,fecfactu from scafac where codclien = " & Cliente & ")"
        If HayRegParaInforme("scafac1", Cad, True) Then
            Cad = "1"
        Else
            Cad = "0"
        End If
    Else
        'ciddirec de scafac
        Cad = DevuelveDesdeBDNew(1, "scaalb", "count(*)", "codclien", Cliente, "N", "", Aux, CStr(codigo), "N")
        
    End If
    If Cad = "" Then Cad = "0"
    If Val(Cad) > 0 Then
        MsgBox "Existen facturas " & Resul, vbExclamation
        Exit Function
    End If
    

    'FALTAR mas cosas
    PuedeEliminarDirecEnvio = True
    
End Function





'Devolvera true SI llega a imprimir
Public Function ComprobarPedidosClientesDesdeAlbProveedor(NumAlbar As String, FechaAlb As Date, Codprove As Long) As Boolean
Dim SQL As String
Dim Cad As String
Dim CAr As Collection
Dim L As Long
Dim J As Integer
Dim ArtiVarios As String

    On Error GoTo EComprobarPedidosClientesDesdeAlbProveedor
    
    
     ComprobarPedidosClientesDesdeAlbProveedor = False
    'Como teinsa NO quiere esto, con "SU parametro" nos salimos sin hacer nada
    If vParamAplic.Frecuencias Then Exit Function
    
    
    
    
    
'Para todas las lineas de un albaran cruzare con la tabla de pedidos de clientes para
'ver que pedidos estan esperando el albran recibido

    Set miRsAux = New ADODB.Recordset
   
    SQL = "Select slialp.codartic,slialp.nomartic,sum(cantidad),sum(artvario) EsDeVarios from slialp,sartic where slialp.codartic=sartic.codartic AND ctrstock = 1"
    SQL = SQL & " AND numalbar = " & DBSet(NumAlbar, "T")
    SQL = SQL & " AND fechaalb = " & DBSet(FechaAlb, "F") & " AND  slialp.codprove = " & Codprove
    SQL = SQL & " GROUP BY 1"
    L = 0
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        'Borro las temporales
        conn.Execute "DELETE from tmpsliped WHERE codusu = " & vUsu.codigo
        conn.Execute "DELETE from tmpnlotes WHERE codusu = " & vUsu.codigo
        conn.Execute "DELETE from tmpslipreu WHERE codusu = " & vUsu.codigo   'articulos varios. Que salgan todas las descripciones
        
        'Meto en tmpnlotes las lineas del albaran con la cantidad total
        Set CAr = New Collection
        Cad = ""
        SQL = ""
        ArtiVarios = ""
        While Not miRsAux.EOF
            L = L + 1
            'tmpnlotes(codusu,numalbar,fechaalb,codprove,numlinea,codartic,nomartic,cantidad)
            SQL = SQL & ", (" & vUsu.codigo & "," & DBSet(NumAlbar, "T") & ","
            SQL = SQL & DBSet(FechaAlb, "F") & "," & Codprove & "," & L & "," & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & "," & DBSet(miRsAux.Fields(2), "N") & " )"
            
            Cad = Cad & ", " & DBSet(miRsAux!codArtic, "T")
            If (L Mod 10) = 0 Then
                Cad = Mid(Cad, 2)
                CAr.Add Cad
                Cad = ""
            End If
            
            
            'Es de varios. Y hay mas de una linea
            'Como cada linea de varios suma 1, tienen que haber mas de uno para que haya dos lineas(o mas )
            ' de un mismo articuilo de varios
            If Val(DBLet(miRsAux!EsDeVarios, "N")) > 1 Then
                'OK. Articulo de varios que hay mas de un articulo
                
                ArtiVarios = ArtiVarios & miRsAux!codArtic & "@#@#" & miRsAux!NomArtic & "|"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        SQL = Mid(SQL, 2)
        SQL = "INSERT tmpnlotes(codusu,numalbar,fechaalb,codprove,numlinea,codartic,nomartic,cantidad) VALUE " & SQL
        conn.Execute SQL
        
        If Cad <> "" Then
            Cad = Mid(Cad, 2)
            CAr.Add Cad
        End If
        
        
        'Los articulos de varios. Por si tienen distinta descripcion
        While ArtiVarios <> ""
            L = InStr(1, ArtiVarios, "|")
            If L = 0 Then
                ArtiVarios = ""
            Else
                Cad = Mid(ArtiVarios, 1, L - 1)
                ArtiVarios = Mid(ArtiVarios, L + 1)
                
                'Separo codartic de nomartic   cod@#@#nom
                L = InStr(1, Cad, "@#@#")  'NO PUE SER 0
                SQL = Mid(Cad, L + 4)  'es el primer nomartic de articulos varios. Para que no ,lo pinte dos veces en el rpt
                Cad = Mid(Cad, 1, L - 1)
                Cad = "WHERE codartic =" & DBSet(Cad, "T")
                Cad = Cad & " AND numalbar = " & DBSet(NumAlbar, "T")
                Cad = Cad & " AND fechaalb = " & DBSet(FechaAlb, "F") & " AND  slialp.codprove = " & Codprove
                Cad = "Select codartic,nomartic from slialp " & Cad
                miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                J = 0
                Cad = SQL
                SQL = ""
                While Not miRsAux.EOF
                    
                    If Cad <> miRsAux!NomArtic Then
                        J = J + 1
                        Cad = miRsAux!NomArtic
                        SQL = SQL & ", (" & vUsu.codigo & ",1,1,1," & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & ")"
                    Else
                        'NO HACEMOS nada
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                
                
                If J > 0 Then
                    SQL = Mid(SQL, 2)
                    Cad = "INSERT INTO tmpslipreu(codusu,numofert,numlinea,codalmac,codartic,nomartic) VALUES " & SQL
                    conn.Execute Cad
                End If
            End If
        Wend
        
        'Ahore veremos en clientes para cada articulo si hay pedido
        'tmpsliped(codusu,fecpedprov,numpedcl,codartic,nomartic,ampliaci,referart,cantidad,numbultos)  'numbultos parar ordenar en el report
        L = 0
        For J = 1 To CAr.Count
            
            'Dic 2013.  Añadimos recogecl,trim(concat(codpobla,' ',pobclien)) para mostrar en el report
                                                                    
            SQL = "select fecpedcl,scaped.numpedcl,sliped.codArtic,NomClien,FecEntre,Cantidad,recogecl,trim(concat(codpobla,' ',pobclien)) lapobla"
            SQL = SQL & ",sliped.codalmac,canstock,artvario"
            SQL = SQL & " from scaped inner join sliped on scaped.numpedcl=sliped.numpedcl  "
            SQL = SQL & " inner join sartic on sliped.codartic=sartic.codartic "
            SQL = SQL & " left join salmac ON sliped.codartic=salmac.codartic and sliped.codalmac=salmac.codalmac"
            SQL = SQL & " where scaped.numpedcl=sliped.numpedcl "
            
            
            
            

            'Añadimos esto
            SQL = SQL & " AND cantidad >0 "
            
            
            
            SQL = SQL & " and sliped.codartic in (" & CAr(J) & ") ORDER BY codartic,fecentre"
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            While Not miRsAux.EOF
                
                ArtiVarios = ""
                'SOLO HERBELCA
                If vParamAplic.NumeroInstalacion = 2 Then
                    'Si el almacen es gandia y castellon NO sale si el stock es cero
                    If miRsAux!codAlmac = 2 Or miRsAux!codAlmac = 4 Then
                        'If miRsAux!CanStock <= 0 Then ArtiVarios = "NO"
                        If miRsAux!CanStock <> 0 Then
                            If miRsAux!artvario = 0 Then ArtiVarios = "NO"
                        End If
                    End If
                End If
                If ArtiVarios = "" Then
                    L = L + 1
                    Cad = Cad & ", (" & vUsu.codigo & ",'" & Format(miRsAux!fecpedcl, "dd/mm/yyyy") & "'," & miRsAux!numpedcl & ","
                    Cad = Cad & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!Nomclien, "T") & "," & DBSet(DBLet(miRsAux!lapobla, "T") & " ", "T") & ",'" & Format(miRsAux!FecEntre, "dd/mm/yyyy") & "',"
                    Cad = Cad & DBSet(miRsAux!cantidad, "N") & "," & L & "," & miRsAux!recogecl & ")"
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If Cad <> "" Then
                Cad = Mid(Cad, 2)
                SQL = "INSERT INTO tmpsliped(codusu,fecpedprov,numpedcl,codartic,nomartic,ampliaci,referart,cantidad,numbultos,numlinea) VALUES " & Cad
                conn.Execute SQL
            End If
        Next
    Else
        miRsAux.Close  'eof
    End If
    
    
    If L > 0 Then
        ComprobarPedidosClientesDesdeAlbProveedor = True
        
        
        
            'Llamaremos a imprimir general
             With frmImprimir
                .FormulaSeleccion = "{tmpnlotes.codusu} = " & vUsu.codigo
                .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
                .NumeroParametros = 1
        
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 5
                .PulsaAceptar = True
                .Titulo = "Listado"
                .NombreRPT = "rpedProvCli.rpt"
                .ConSubInforme = False
                .Show vbModal
            End With
    
    
      
        
        
    End If

EComprobarPedidosClientesDesdeAlbProveedor:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
    Set CAr = Nothing
End Function





'INSERTA lineas tb, con un frame
Public Function InsertarDesdeForm2(ByRef Formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In Formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.Value = 1 Then
                        Cad = "1"
                        Else
                        Cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                    Der = Der & Cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                        If Control.ListIndex = -1 Then
                            Cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.columna & ""
                            Cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
            
        'ElseIf TypeOf Control Is DTPicker Then
        ElseIf False Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
''                        If Control.Value Then
''                            If Izda <> "" Then Izda = Izda & ","
''                            Izda = Izda & "" & mTag.columna & ""
''                            cad = Control.index
''                            If Der <> "" Then Der = Der & ","
''                            Der = Der & cad
''                        End If
'                        If Izda <> "" Then Izda = Izda & ","
'                        Izda = Izda & "" & mTag.columna & ""
'
'                        'Parte VALUES
'                        If Control.visible Then
'                            Cad = ValorParaSQL(Control.Value, mTag)
'                        Else
'                            Cad = ValorNulo
'                        End If
'                        If Der <> "" Then Der = Der & ","
'                        Der = Der & Cad
'                    End If
'                End If
'            End If
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
       
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function





'******************************************************************************
'******************************************************************************
'
Public Function ComprobarTotalPendienteFormasPagoRecFinan(Cliente As Long, ForpasRecFinan As String, ImporteVenta As Currency) As Boolean
Dim R As ADODB.Recordset
Dim SQL As String
Dim Cta As String
Dim Importe As Currency
Dim Aux As Currency
Dim ImporteCredito As Currency


    ComprobarTotalPendienteFormasPagoRecFinan = True
    
    Set R = New ADODB.Recordset
    SQL = "Select codmacta,limcredi from sclien where codclien=" & Cliente
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO puede ser eof
    Cta = DBLet(R.Fields(0), "T")
    ImporteCredito = DBLet(R!limcredi, "N")
    R.Close
        
    Importe = 0
        
    'VEMOS EN TESORERIA
    SQL = "Select sum(impvenci)+sum(coalesce(gastos,0))-sum(coalesce(impcobro,0)) from "
    SQL = SQL & IIf(vParamAplic.ContabilidadNueva, "cobros", "scobro")
    SQL = SQL & " WHERE codmacta = '" & Cta & "' AND codforpa IN " & ForpasRecFinan
    R.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then Importe = Importe + DBLet(R.Fields(0), "N")
    R.Close
    
    
    'Los albaranes
    SQL = "select codtipom,numalbar from scaalb where codclien=" & Cliente & " and codforpa in " & ForpasRecFinan
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cta = ""
    While Not R.EOF
        Cta = Cta & ", (" & DBSet(R!codtipom, "T") & "," & R!NumAlbar & ")"
        R.MoveNext
    Wend
    R.Close
    
    
    'Si hay albaranes
    If Cta <> "" Then
        Cta = Mid(Cta, 2)
        Cta = "where (codtipom, numalbar) in (" & Cta & ") AND "
        Cta = Cta & "slialb.codartic=sartic.codartic group by 1"
        
        Cta = "select codigiva,sum(importel) from slialb ,sartic " & Cta

        R.Open Cta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not R.EOF
            Cta = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", R.Fields(0))
            If Cta = "" Then Cta = "0"
            Aux = CCur(Cta)
            Aux = (Aux / 100) + 1
            Aux = DBLet(R.Fields(1), "N") * Aux
            Importe = Importe + Round(Aux, 2)
            R.MoveNext
        Wend
        R.Close
        
        
    End If

    Importe = Importe + ImporteVenta
    
    'Tengo importe e importe credito
    If ImporteCredito < Importe Then
        Importe = Importe - ImporteCredito
        MsgBox "El importe excede del maximo permitido en " & Format(Importe, FormatoImporte), vbExclamation
        ComprobarTotalPendienteFormasPagoRecFinan = False
    End If

    Set R = Nothing
End Function
