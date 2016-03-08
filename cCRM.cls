VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'Datos importantes
Public CodClien As Long
Public codmacta As String


'Para saber que datos va a mostrar

















Public Sub BorrarTemporales()
    'tmpcrmclien tmpcrmcobros tmpcrmmsg tmpcrmtesor
    conn.Execute "DELETE FROM tmpcrmclien WHERE codusu = " & vUsu.codigo
    conn.Execute "DELETE FROM tmpcrmcobros WHERE codusu = " & vUsu.codigo
    conn.Execute "DELETE FROM tmpcrmmsg WHERE codusu = " & vUsu.codigo
    conn.Execute "DELETE FROM tmpcrmtesor WHERE codusu = " & vUsu.codigo
    
End Sub





Private Sub DatosFacturacion()
    
End Sub
