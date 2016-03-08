VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad

Private mvarNombre As String
Private mvarTipoDato As String
Private mvarVacio As String
Private mvarMaximo As Long
Private mvarMinimo As Long
Private mvarTabla As String
Private mvarColumna As String
Private mvarFormato As String
Private mvarCargado As Boolean 'Indicará si el objeto está cargado o no.
Private mvarEsClave As Boolean  ' Indica si es campo clave en la BD
Private Mens As String

Private vValores As String  'Variable sobre la que se recoge el TAG en CARGAR
Private vValor              'Variable sobre la que se recoge el .Text del objeto en COMPROBAR
Private vValorF             'Variable sobre la que se recoge el .Text del objeto en DARFORMATO
Private VacioMin As Boolean 'True significa que mvarMinimo esta vacio y False que no lo está
Private VacioMax As Boolean 'True significa que mvarMaximo esta vacio y False que no lo está

'---- Añade: Laura 10/06/05
Dim valorDec As Currency
'-------------------------------


'------------------------------------------------------------------------
'                   Descripción del Objeto CTag
'------------------------------------------------------------------------
    '
    '   V1|V2|V3|V4|V5|V6|V7|V8|V9| (Es necesario que la secuencia termine con |)
    '   V1: Nombre del campo, para mensajes. No puede ir vacío.
    '   V2: Tipo de campo N=Numérico,T=Texto,F=Fecha(dd/mm/aaaa),H=Hora(hh:mm). No puede ir vacío.
    '   V3: S=Se permiten nulos, N=No se permiten. No puede ir vacío.
    '   V4: El valor ha de ser mayor o igual que él. Puede ir vacío.
    '   V5: El valor a de ser menor o igual que él. Puede ir vacío.
    '   V6: Nombre de la tabla que soporta el dato. Puede ir vacío.
    '   V7: Nombre de la columna que soporta el dato. Puede ir vacío.
    '   V8: Formato de la caja (Máscara). Puede ir vacío.
    '   V9: Para saber si el campo es clave en la BD
    'Si V4 o V5 van vacíos no se hace comprobación de rangos.
    'Si el dato proviene de un base de datos,utilizaremos V6 y V7.
    'Si no, V6 y V7 irán vacíos.
    '------------------------------------------------------------------------
    
    
    
'----------------------------------------------------------------------------
'                   Propiedades del modulo CTAG
'----------------------------------------------------------------------------
Public Property Get Cargado() As Boolean
'Propiedad interna.
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Cargado
    Cargado = mvarCargado
End Property



Public Property Get EsClave() As Boolean
'Propiedad interna.
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Cargado
    EsClave = mvarEsClave
End Property


Public Property Get Nombre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property

Public Property Let Nombre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property

Public Property Get TipoDato() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.TipoDato
    TipoDato = mvarTipoDato
End Property

Public Property Let TipoDato(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.TipoDato = 5
    mvarTipoDato = vData
End Property

Public Property Get Vacio() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Vacio
    Vacio = mvarVacio
End Property

Public Property Let Vacio(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Vacio = 5
    mvarVacio = vData
End Property

Public Property Get Minimo() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Minimo
    Minimo = mvarMinimo
End Property

Public Property Let Minimo(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Minimo = 5
    mvarMinimo = vData
End Property

Public Property Get Maximo() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Maximo
    Maximo = mvarMaximo
End Property

Public Property Let Maximo(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Maximo = 5
    mvarMaximo = vData
End Property

Public Property Get tabla() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Tabla
    tabla = mvarTabla
End Property

Public Property Let tabla(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Tabla = 5
    mvarTabla = vData
End Property

Public Property Get columna() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Columna
    columna = mvarColumna
End Property

Public Property Let columna(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Columna = 5
    mvarColumna = vData
End Property

Public Property Get Formato() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Formato
    Formato = mvarFormato
End Property

Public Property Let Formato(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Formato = 5
    mvarFormato = vData
End Property




'-------------------------------------------------------------------------
'                   FUNCIONES del modulo CTag (Métodos)
'-------------------------------------------------------------------------



'****************************   CARGAR  **********************************

'Asigna cada uno de los valores que recibe (en forma string->V1|V2|...), a las propiedades correspondientes.
'Devuelve True si la asignación se ha hecho con exito.
'Devuelve False si la asignación ha fallado, por ejemplo porque se le han pasado
'más valores que campos existen.

Public Function Cargar(vObjetoCarga As Object) As Boolean
    
    Dim V(9) As Variant
    Dim i As Integer
    Dim i2 As Integer
    Dim mCadena As String
    Dim Mc As String
    
    
    vValores = vObjetoCarga.Tag
    
    If vValores = "" Then
        '---- Modifica:Laura
        Cargar = True
        mvarCargado = False
        '-----------------------------
        Exit Function
    Else
    
        VacioMin = True
        VacioMax = True
        
        '-- Limpiamos los campos
        For i = 1 To 9
            V(i) = ""
        Next i
        '-- Cargamos los datos
        mCadena = ""
        Mc = ""
        i2 = 0
        i = 1
        While (i <= Len(vValores)) And (i2 < 10)
            Mc = Mid(vValores, i, 1)
            If Mc = "|" Then
                i2 = i2 + 1
                V(i2) = mCadena
                mCadena = ""
            Else
                mCadena = mCadena & Mc
            End If
            i = i + 1
        Wend
        If i2 <> 9 Then
        'Se han introducido más parámetros de los que se admiten
            
            If i2 = 0 Then
                Cargar = True
            Else
                Cargar = False
            End If
            mvarCargado = False
        Else
            Cargar = False 'Por defecto
            mvarCargado = False
            '-- Comprobamos que no se han dejado ningún campo de los obligatorios
            If V(1) = "" Then
                MsgBox "El Nombre del campo no puede estar vacío"
                Cargar = False
                Exit Function
            End If
            If V(2) = "" Then
                MsgBox "Debe indicar un tipo para el campo"
                Cargar = False
                Exit Function
            Else
                If (V(2) <> "T") And (V(2) <> "F") And (V(2) <> "N") And (V(2) <> "H") Then
                    MsgBox "Tipo de Campo desconocido"
                    Cargar = False
                    Exit Function
                End If
            End If
            If V(3) = "" Then
                MsgBox "Debe indicar si el campo puede estar vacío o no"
                Cargar = False
                Exit Function
            Else
                If (V(3) <> "N") And (V(3) <> "S") Then
                    MsgBox "Debe indicar si el campo puede estar vacio, se acepta S o N"
                    Cargar = False
                    Exit Function
                End If
            End If
            If V(4) <> "" Then
                VacioMin = False
                If Not IsNumeric(V(4)) Then
                    MsgBox "El Valor Mínimo ha de ser numérico"
                    Cargar = False
                    Exit Function
                End If
            End If
            If V(5) <> "" Then
                VacioMax = False
                If Not IsNumeric(V(5)) Then
                    MsgBox "El Valor Máximo ha de ser numérico"
                    Cargar = False
                    Exit Function
                End If
            End If
            
            mvarNombre = V(1)
            mvarTipoDato = V(2)
            mvarVacio = V(3)
            mvarMinimo = Val(V(4))
            mvarMaximo = Val(V(5))
            mvarTabla = V(6)
            mvarColumna = V(7)
            mvarFormato = V(8)
            mvarEsClave = (V(9) = "S")
            mvarCargado = True
            Cargar = True
        End If
    End If
End Function

'****************************   COMPROBAR ********************************

'Comprueba que la cadena que se le pasa cumple lo especificado en cada una de las propiedades.
Public Function Comprobar(vObjetoComp As Object) As Boolean
   
    
    If mvarCargado = False Then
        Call Cargar(vObjetoComp)
    End If
    If vValores = "" Then
        Comprobar = True
        Exit Function
    Else
        vValor = vObjetoComp.Text
        'Comprobamos si permite nulos
        If mvarVacio = "N" Then
            '-- No se permiten nulos.
            If vValor = "" Then
                Mens = "El valor de " & mvarNombre & " no puede ser nulo."
                MsgBox Mens, vbInformation, "Comprobador de campos"
                Comprobar = False
                Exit Function
            End If
        Else
            'MODIFICACION DEL 1 DE ABRIL de 2003
            'Si se permite vacio, y esta vacio, SALIMOS
            If vValor = "" Then
                Comprobar = True
                Exit Function
            End If
        End If
    
        Select Case mvarTipoDato
            Case "N"
                If vValor <> "" Then
                    If Not IsNumeric(vValor) Then
                        Mens = "El valor de " & mvarNombre & " debe ser numérico."
                        MsgBox Mens, vbInformation, "Comprobador de campos"
                        Comprobar = False
                        Exit Function
                    End If
                End If
            Case "T"
                'Se interpreta que si no es de otro tipo, es de texto.
                
            Case "F"
                    If vValor <> "" Then
                        If Not EsFechaOK(CStr(vValor)) Then
                            Mens = "El valor de " & mvarNombre & " debe ser una fecha (dd/mm/aaaa)."
                            MsgBox Mens, vbInformation, "Comprobador de campos"
                            Comprobar = False
                            Exit Function
                        End If
                    End If
            Case "H"
                If Not IsDate(vValor) Then
                    Mens = "El valor de " & mvarNombre & " debe ser una hora (hh:mm)."
                    MsgBox Mens, vbInformation, "Comprobador de campos"
                    Comprobar = False
                    Exit Function
                End If
        End Select
        
        'Si es numérico enonces
        If mvarTipoDato = "N" Then
        'Si hay un valor mínimo con el que comparar.
            If VacioMin = False Then
                  '---- Modifica: Laura 10/06/05
                  valorDec = ImporteFormateado(CStr(vValor))
                  If Val(valorDec) < mvarMinimo Then
                  'ANTES: If Val(vValor) < mvarMinimo Then
                  '------------------------------------------
                      Mens = "El valor de " & mvarNombre & " debe ser mayor o igual que " & mvarMinimo
                      MsgBox Mens, vbInformation, "Comprobador de campos"
                      Comprobar = False
                      Exit Function
                  End If
            End If
            
        'Si hay un valor máximo con el que comparar.
            If VacioMax = False Then
                  '---- Modifica: Laura 10/06/05
                  'ANTES: If Val(vValor) > mvarMaximo Then
                  If Val(valorDec) > mvarMaximo Then
                  '----------------------------------------
                      Mens = "El valor de " & mvarNombre & " debe ser menor o igual que " & mvarMaximo
                      MsgBox Mens, vbInformation, "Comprobador de campos"
                      Comprobar = False
                      Exit Function
                  End If
            End If
        End If
        Comprobar = True
    End If

End Function

Public Sub DarFormato(ObjetoFormat As Object)
    
    'Si el objeto no está cargado, lo cargamos.
    If mvarCargado = False Then
        Call Cargar(ObjetoFormat)
    End If
    'Si el objeto se ha cargado correctamente, formateamos.
    If mvarCargado = True Then
        vValorF = ObjetoFormat.Text
        On Error GoTo Error_DarFormato
            If (mvarFormato <> "") And (vValorF <> "") Then
                If mvarTipoDato = "F" Or mvarTipoDato = "H" Then
                    'Si el Tipo de Dato es Hora o Fecha, IsDate(Dato) tiene que ser True.
                    If EsFechaOK(CStr(vValorF)) Then
                        'vValorF = Format(vValorF, mvarFormato)
                    Else
                        MsgBox "El dato introducido no es una hora o fecha valida"
                    End If
                Else
                    'Si el dato no tiene que ser una hora ni una fecha, formateamos, sin más.
                    vValorF = Format(vValorF, mvarFormato)
                End If
            End If
    Else
        MsgBox "No se puede dar formato al objeto ya que no ha sido cargado correctamente"
        Exit Sub
    End If
    ObjetoFormat.Text = vValorF
Error_DarFormato: 'Control de errores producidos al dar formato
    If Err.Number = 6 Then
        MsgBox "Compruebe que el dato introducido sea coherente con la máscara"
        Exit Sub
    End If
End Sub

Private Sub Class_Initialize()
    'Al inicializar el objeto, no está cargado y por eso mvarCargado se inicializa a False
    mvarCargado = False
End Sub
