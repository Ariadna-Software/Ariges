VERSION 5.00
Begin VB.Form frmParaBorrar 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AJustar smoval"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
End
Attribute VB_Name = "frmParaBorrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim Cad As String
Dim RA As ADODB.Recordset
Dim RArt As ADODB.Recordset
Dim I As Integer
Dim J As Integer


    Cad = "select * from smoval where detamovi in ('ALZ','ALV') and observa like 'N/%'"
  
    Set RA = New ADODB.Recordset
    Set RArt = New ADODB.Recordset
    
    RArt.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RArt.EOF
        Label1.Caption = RArt!codArtic & " " & RArt!FechaMov
    
        miSQL = RArt!document
        I = InStr(1, miSQL, "/")
        'NO puede ser 0
        If I > 0 Then
           
            miSQL = Trim(Mid(miSQL, I + 1))
            I = InStr(miSQL, "(")
            If I > 0 Then
                miSQL = Trim(Mid(miSQL, 1, I - 1))
                
            End If
        End If
        
        If I > 0 Then
            Cad = Format(Val(miSQL), "00000")
            If RArt!FechaMov < CDate("01/01/2019") Then Cad = "18" & Cad
        
            miSQL = "Select * from slifac where codartic=" & DBSet(RArt!codArtic, "T") & " AND codtipoa=" & DBSet(RArt!DetaMovi, "T")
            miSQL = miSQL & " AND numalbar =" & DBSet(Cad, "T")
    
    
    
            RA.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RA.EOF Then
     
            Else
                RA.Close
                
                miSQL = "Select * from slialb where codartic=" & DBSet(RArt!codArtic, "T") & " AND codtipom=" & DBSet(RArt!DetaMovi, "T")
                miSQL = miSQL & " AND numalbar =" & Cad
                RA.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RA.EOF Then
                
                
                
                Else
                    Cad = ""
                End If
            End If
            If Cad <> "" Then
                
                miSQL = "UPDATE smoval set numlinea= " & RA!numlinea & ", document = " & DBSet(Cad, "T")
                miSQL = miSQL & " WHERE codartic =" & DBSet(RArt!codArtic, "T")
                miSQL = miSQL & " AND fechamov =" & DBSet(RArt!FechaMov, "F")
                miSQL = miSQL & " AND tipomovi =" & DBSet(RArt!tipomovi, "T")
                miSQL = miSQL & " AND detamovi =" & DBSet(RArt!DetaMovi, "T")
                miSQL = miSQL & " AND document =" & DBSet(RArt!document, "T")
                miSQL = miSQL & " AND cantidad=" & DBSet(RArt!cantidad, "N")
                conn.Execute miSQL
            Else
                
                
                
'                If RArt!FechaMov > CDate("01/02/2019") Then MsgBox RArt!document
                
            End If
            
            
            RA.Close

        Else
 '           MsgBox RArt!document
        End If
    
        


        RArt.MoveNext
    Wend
    RArt.Close
    




End Sub

Private Sub Command2_Click()


   '1519 04159
'1513   04658
 


AbrirConsultaPrecio2 1746, "TU0004", "", "ROTONDA ALCOY"
End Sub
