Attribute VB_Name = "ModBuscaGrid"
Option Explicit
'=======================================================
'Este modulo utiliza funciones del modulo: ModFunciones
'=======================================================



Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim cad As String
'====Modificado por Laura Junio 2004:
'====Se añade el formato empipado
'Montamos al final: "Cod Diag.|tabla|columna|tipo|formato|10·"

    ParaGrid = ""
    cad = ""
    Set mTag = New CTag
    mTag.Cargar Control
    If mTag.Cargado Then
        If Control.Tag <> "" Then
            'Si es texto monta esta parte de sql
            If TypeOf Control Is TextBox Then
                If Desc <> "" Then
                    cad = Desc
                Else
                    cad = mTag.Nombre
                End If
                cad = cad & "|"
                
                '----------------------
                'Añade Laura - 1/9/04
                cad = cad & mTag.Tabla & "|"
                '----------------------
                
                cad = cad & mTag.columna & "|"
                cad = cad & mTag.TipoDato & "|"
                
                '----------------------
                'Añade Laura - Junio/04
                cad = cad & mTag.Formato & "|"
                '----------------------
                
                cad = cad & AnchoPorcentaje & "·"
    
            'CheckBOX
            ElseIf TypeOf Control Is CheckBox Then
    
            ElseIf TypeOf Control Is ComboBox Then
                If Desc <> "" Then
                    cad = Desc
                Else
                    cad = mTag.Nombre
                End If
                cad = cad & "|"
                '----------------------
                'Añade Laura - 1/9/04
                cad = cad & mTag.Tabla & "|"
                '----------------------
                cad = cad & mTag.columna & "|"
                cad = cad & mTag.TipoDato & "|"
                cad = cad & mTag.Formato & "|"
                cad = cad & AnchoPorcentaje & "·"
            
    
            End If 'De los elseif
        End If
        Set mTag = Nothing
        ParaGrid = cad
    End If
End Function




''////////////////////////////////////////////////////
'' Monta a partir de una cadena devuelta por el formulario
''de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As CTag
Dim cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

    ValorDevueltoFormGrid = ""
    cad = ""
    Set mTag = New CTag
    mTag.Cargar Control
    If mTag.Cargado Then
        If Control.Tag <> "" Then
            'Si es texto monta esta parte de sql
            If TypeOf Control Is TextBox Then
                Aux = RecuperaValor(CadenaDevuelta, Orden)
                If Aux <> "" Then cad = mTag.columna & " = " & ValorParaSQL(Aux, mTag)
            'CheckBOX
           ' ElseIf TypeOf Control Is CheckBox Then
           '
           ' ElseIf TypeOf Control Is ComboBox Then
           '
           '
            End If 'De los elseif
        End If
    End If
    Set mTag = Nothing
    ValorDevueltoFormGrid = cad
End Function




