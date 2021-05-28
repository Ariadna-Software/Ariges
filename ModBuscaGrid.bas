Attribute VB_Name = "ModBuscaGrid"
Option Explicit
'=======================================================
'Este modulo utiliza funciones del modulo: ModFunciones
'=======================================================



Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As cTag
Dim Cad As String
'====Modificado por Laura Junio 2004:
'====Se añade el formato empipado
'Montamos al final: "Cod Diag.|tabla|columna|tipo|formato|10·"

    ParaGrid = ""
    Cad = ""
    Set mTag = New cTag
    mTag.Cargar Control
    If mTag.Cargado Then
        If Control.Tag <> "" Then
            'Si es texto monta esta parte de sql
            If TypeOf Control Is TextBox Then
                If Desc <> "" Then
                    Cad = Desc
                Else
                    Cad = mTag.Nombre
                End If
                Cad = Cad & "|"
                
                '----------------------
                'Añade Laura - 1/9/04
                Cad = Cad & mTag.tabla & "|"
                '----------------------
                
                Cad = Cad & mTag.columna & "|"
                Cad = Cad & mTag.TipoDato & "|"
                
                '----------------------
                'Añade Laura - Junio/04
                Cad = Cad & mTag.Formato & "|"
                '----------------------
                
                Cad = Cad & AnchoPorcentaje & "·"
    
            'CheckBOX
            ElseIf TypeOf Control Is CheckBox Then
    
            ElseIf TypeOf Control Is ComboBox Then
                If Desc <> "" Then
                    Cad = Desc
                Else
                    Cad = mTag.Nombre
                End If
                Cad = Cad & "|"
                '----------------------
                'Añade Laura - 1/9/04
                Cad = Cad & mTag.tabla & "|"
                '----------------------
                Cad = Cad & mTag.columna & "|"
                Cad = Cad & mTag.TipoDato & "|"
                Cad = Cad & mTag.Formato & "|"
                Cad = Cad & AnchoPorcentaje & "·"
            
    
            End If 'De los elseif
        End If
        Set mTag = Nothing
        ParaGrid = Cad
    End If
End Function




''////////////////////////////////////////////////////
'' Monta a partir de una cadena devuelta por el formulario
''de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As cTag
Dim Cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

    ValorDevueltoFormGrid = ""
    Cad = ""
    Set mTag = New cTag
    mTag.Cargar Control
    If mTag.Cargado Then
        If Control.Tag <> "" Then
            'Si es texto monta esta parte de sql
            If TypeOf Control Is TextBox Then
                Aux = RecuperaValor(CadenaDevuelta, Orden)
                If Aux <> "" Then Cad = mTag.columna & " = " & ValorParaSQL(Aux, mTag)
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
    ValorDevueltoFormGrid = Cad
End Function




