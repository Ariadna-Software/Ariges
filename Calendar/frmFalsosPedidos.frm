VERSION 5.00
Begin VB.Form frmComEntPedidos 
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmComEntPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public DatosADevolverBusqueda2 As String
Public EsHistorico As Boolean
