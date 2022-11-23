VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTPVCuadreCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text1"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   14
      Left            =   8400
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   14
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton cmdImprimeFraCli 
      Height          =   495
      Left            =   7560
      Picture         =   "frmFacTPVCuadreCaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Imprimir"
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   9240
      TabIndex        =   67
      Tag             =   "Final|N|N|0||||#,##0.00||"
      Top             =   3720
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   65
      Tag             =   "Final|N|N|0||||#,##0.00||"
      Top             =   3720
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   62
      Tag             =   "Final|N|N|0||||#,##0.00||"
      Top             =   3720
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   59
      Tag             =   "Final|N|N|0||||#,##0.00||"
      Top             =   3720
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   57
      Tag             =   "Final|N|N|0||||#,##0.00||"
      Top             =   3720
      Width           =   1380
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   8400
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   8400
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   6480
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   8400
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6480
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   6000
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   8400
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6000
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   5520
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   8400
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5520
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   5040
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   8400
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5040
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   4560
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   8400
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4560
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   7440
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   2760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   7440
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   6960
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   2760
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6960
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   6480
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   2760
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6480
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   6000
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2760
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6000
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   5520
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5520
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   5040
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   5040
      Width           =   1020
   End
   Begin VB.TextBox txtCuadreImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   4560
      Width           =   1380
   End
   Begin VB.TextBox txtCuadre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4560
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      Tag             =   "Final|N|N|0||||#,##0.00||"
      Top             =   3720
      Width           =   1380
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   2040
      Width           =   2955
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Top             =   3120
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   1440
      Width           =   4275
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   840
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Tag             =   "Inicial|N|N|0||||#,##0.00||"
      Text            =   "Text1"
      Top             =   2640
      Width           =   1380
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   6480
      TabIndex        =   26
      Top             =   120
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblCuadre 
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   5880
      TabIndex        =   72
      Top             =   7980
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio caja sig"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   9240
      TabIndex        =   68
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblCuadreTit 
      Caption         =   "Cuadre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   28
      Top             =   4080
      Width           =   1665
   End
   Begin VB.Line Line1 
      X1              =   10800
      X2              =   120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblIndic 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Leyendo BBDD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   240
      TabIndex        =   66
      Top             =   8160
      Width           =   4725
   End
   Begin VB.Label Label1 
      Caption         =   "Efectivo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   240
      TabIndex        =   64
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Terminal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   3000
      TabIndex        =   63
      Top             =   2690
      Width           =   3225
   End
   Begin VB.Label Label1 
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   3360
      TabIndex        =   61
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Diferencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   7080
      TabIndex        =   60
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Retirada "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   1800
      TabIndex        =   58
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblCuadre 
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   5880
      TabIndex        =   55
      Top             =   7500
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   5880
      TabIndex        =   53
      Top             =   7020
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   5880
      TabIndex        =   51
      Top             =   6540
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   5880
      TabIndex        =   49
      Top             =   6060
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   5880
      TabIndex        =   47
      Top             =   5580
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   5880
      TabIndex        =   45
      Top             =   5100
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   5880
      TabIndex        =   43
      Top             =   4620
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   41
      Top             =   7500
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   39
      Top             =   7020
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   37
      Top             =   6540
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   35
      Top             =   6060
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   33
      Top             =   5580
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   31
      Top             =   5100
      Width           =   2505
   End
   Begin VB.Label lblCuadre 
      Caption         =   "Cincuenta céntimos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   4620
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "€ caja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   5520
      TabIndex        =   27
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Terminal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   24
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   1440
      Width           =   1725
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Saldo inicial"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmFacTPVCuadreCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Fecha As Date
Public abrir As Boolean   'vs cerrar
Public ForzarApertura As Boolean    'Si o si tenemos que abrir caja. Será la primera . No hay ninguna abierta

Dim Cad As String


Private Sub cmdAceptar_Click()
    Dim OK As Boolean
    If abrir Then
        If Text1(0).Text = "" Then Exit Sub
        
        If Not EsNumerico(Text1(0).Text) Then Exit Sub
    
        Cad = "numtermi = " & vParamTPV.NumeroDeTerminal & " AND diacierre is null AND 1"
        Cad = DevuelveDesdeBD(conAri, "diaapertura", "stpvdiacaja", Cad, "1")
        If Cad <> "" Then
            MsgBox "Caja abierta todavia: " & Cad, vbExclamation
            Exit Sub
        End If
        
        
        Cad = "Va a aperturar la caja . " & vbCrLf
        Cad = Cad & "Fecha: " & Format(Fecha, "dd/mm/yyyy") & "       Importe: " & Text1(0).Text & vbCrLf
        Cad = Cad & vbCrLf & "¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        Cad = DevuelveDesdeBD(conAri, "max(id)", "stpvdiacaja", "1", "1")
        Cad = Val(Cad) + 1
        
        
        
        
        
        Cad = Cad & "," & vParamTPV.NumeroDeTerminal & "," & DBSet(Fecha, "F") & "," & DBSet(Now, "FH") & ","
        Cad = Cad & DBSet(Text1(0).Text, "N") & "," & Text2(1).Tag & ")"
        Cad = "INSERT INTO stpvdiacaja(id,numtermi,fecha,diaapertura,inicial,codtrabaAper) VALUES (" & Cad
        If Not ejecutar(Cad, False) Then Exit Sub
        
        
    Else
        If Text1(1).Text = "" Then
            MsgBox "Importe cierre !!!!", vbExclamation
            Exit Sub
        End If
        
        If Not EsNumerico(Text1(1).Text) Then Exit Sub
    
    
        'Imprimimos
        Imprimir
    
        
    
        'Va a cerrar la caja.
        Cad = ""
        If Text1(6).Text <> "" Then
            If Not EsNumerico(Text1(6).Text) Then
                Exit Sub
            Else
                Cad = vbCrLf & "PROXIMA APERTURA: " & Format(DateAdd("d", 1, Now), "dd/mm/yyyy") & vbCrLf
                Cad = Cad & "Importe apertura: " & Text1(6).Text & vbCrLf
            End If
        End If
    
        Cad = "Importe cierre: " & Text1(1).Text & vbCrLf & Cad
        Cad = String(40, "*") & vbCrLf & vbCrLf & Cad & vbCrLf & "¿Continuar?" & vbCrLf & vbCrLf & String(40, "*")
        If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then Exit Sub
        
        
    
        conn.BeginTrans
        OK = False
    
        Cad = "UPDATE stpvdiacaja SET diacierre =" & DBSet(Now, "FH") & ", final ="
        Cad = Cad & DBSet(Text1(1).Text, "N") & ",codtrabaCierr = " & Text2(1).Tag
        Cad = Cad & " WHERE id=" & Text2(0).Tag
        If ejecutar(Cad, False) Then
            
            'Las monedas de cierre
            Cad = ""
            For NumRegElim = 0 To Me.txtCuadre.Count - 1
                If Me.txtCuadre(NumRegElim).visible Then
                    If txtCuadre(NumRegElim).Text <> "" Then
                        'id,tipomoneda,uds,importe
                        Cad = Cad & ",  (" & Text2(0).Tag & ","
                        Cad = Cad & DBSet(lblCuadre(NumRegElim).Tag, "N") & "," & DBSet(txtCuadre(NumRegElim).Text, "N") & "," & DBSet(txtCuadreImporte(NumRegElim).Text, "N") & ")"
                    End If
                End If
            Next
            
            If Cad <> "" Then
                Cad = Mid(Cad, 2)
                Cad = "INSERT INTO stpvdiacaja_cuadre(id,tipomoneda,uds,importe)      VALUES " & Cad
                If Not ejecutar(Cad, True) Then
                    MsgBox "Error insertando cuadre", vbExclamation
                    Text1(6).Text = "NOO"
                End If
            End If
            
            
            If Text1(6).Text = "" Then
                OK = True
                
            Else
                'Aperturamos dia siguiente
                Cad = DevuelveDesdeBD(conAri, "max(id)", "stpvdiacaja", "1", "1")
                Cad = Val(Cad) + 1
                Cad = Cad & "," & vParamTPV.NumeroDeTerminal & "," & DBSet(Now + 1, "F") & ",'" & Format(Now + 1, FormatoFecha) & " " & Format(Now, "hh:mm:ss") & "',"
                Cad = Cad & DBSet(Text1(6).Text, "N") & "," & Text2(1).Tag & ")"
                Cad = "INSERT INTO stpvdiacaja(id,numtermi,fecha,diaapertura,inicial,codtrabaAper) VALUES (" & Cad
                If ejecutar(Cad, False) Then OK = True
                
            End If
        End If
            
            If OK Then
                conn.CommitTrans
            Else
                conn.RollbackTrans
            End If
            If Not OK Then Exit Sub
    End If
    
    CadenaDesdeOtroForm = "OK"
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdImprimeFraCli_Click()
    If abrir Then Exit Sub
    
    
    Imprimir
    
    
End Sub

Private Sub Form_Activate()
    If cmdAceptar.Tag = 1 Then
        cmdAceptar.Tag = 0
        
        'Cargamos datos ventas dia
        Screen.MousePointer = vbHourglass
        If Not abrir Then CargarDatosVentaDia
        Me.lblIndic.visible = False
        cmdAceptar.visible = True
        
        PonerFoco Me.Text1(1)
        
        DoEvents
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    limpiar Me
    Label1(9).Caption = "" 'por si el cierre no lo hace el mismo trabajador
    Label1(5).visible = Not abrir
    Label1(5).visible = Not abrir
    Text1(1).visible = Not abrir
    ListView1.visible = Not abrir
    
    Text1(0).Enabled = abrir
    
    cmdAceptar.Left = 3000
    cmdCancelar.Left = 4500
    
    Caption = "Caja"
    If Not abrir Then
        Text1(0).Locked = True
        cmdAceptar.Left = 8280
        cmdCancelar.Left = 9720
        Me.Width = 11145
        
        Label7.Caption = "Cierre caja"
                
        cmdAceptar.Left = 8280
        cmdCancelar.Left = 9720
    Else
        
        Label7.Caption = "Apertura caja"
        Text1(0).Locked = False
        cmdAceptar.Left = 2520
        cmdCancelar.Left = 4080
        Me.Width = 6000
        Me.Height = 3800
    End If
    Line1.visible = Not abrir
    Me.cmdImprimeFraCli.visible = Not abrir
    For NumRegElim = 1 To 6
        Me.Text1(NumRegElim).visible = Not abrir
    Next
    For NumRegElim = 4 To 10
        Me.Label1(NumRegElim).visible = Not abrir
    Next
    For NumRegElim = 0 To txtCuadre.Count - 1
        txtCuadre(NumRegElim).Locked = abrir
    Next
    lblCuadreTit.visible = Not abrir
    
    
    Text2(0).Text = Format(Fecha, "dd/mm/yyyy")
    Text2(1).Tag = PonerTrabajadorConectado(Cad)
    If Cad = "" Then Cad = "Error trabajador conectado"
    Text2(1).Text = Text2(1).Tag & " - " & Cad
    
    'destermi   numtermi   spatpvt
    Cad = DevuelveDesdeBD(conAri, "destermi", "spatpvt", "numtermi", vParamTPV.NumeroDeTerminal)
    Text2(2).Text = Cad
    
    If abrir Then
        Text1(0).Text = "0,00"
    Else
       ' Text1(0).Text = Format(DBLet(miRsAux!Inicial, "N"), FormatoCantidad)
    End If
    Text1(1).Text = ""
    cmdAceptar.Tag = 1
    
    CuadreVisible Not abrir
    Screen.MousePointer = vbHourglass
    
    
    
    If abrir Then
        '4440
        cmdAceptar.Top = 3440
        
        
    Else
        If Me.txtCuadre(13).visible Then
            NumRegElim = txtCuadre(13).Top
        ElseIf Me.txtCuadre(12).visible Then
            NumRegElim = txtCuadre(12).Top
        Else
            NumRegElim = txtCuadre(11).Top
        End If
        NumRegElim = NumRegElim + txtCuadre(11).Height + 540
        cmdAceptar.Top = NumRegElim
        Me.cmdImprimeFraCli.Top = cmdAceptar.Top
    End If
    cmdCancelar.Top = cmdAceptar.Top
    Me.Height = cmdAceptar.Top + cmdAceptar.Height + 540
    
    
    
    
    cmdAceptar.visible = False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
      KEYpress KeyAscii

End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


'++
'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim mTag As cTag
    
    'If Not PerderFocoGnral(Text1(Index), 3) Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0, 1, 6
                Text1(Index).Text = Trim(Text1(Index).Text)
                If Text1(Index).Text = "" Then Exit Sub
                Set mTag = New cTag
                If mTag.Cargar(Text1(Index)) Then
                    If mTag.Cargado Then
                        If mTag.Comprobar(Text1(Index)) Then
                            If Not PonerFormatoDecimal(Text1(Index), 3) Then Text1(Index).Text = ""
                        Else
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                End If
                Set mTag = Nothing

        
    End Select
    If Not abrir Then ImporteFinal
    '---
End Sub



Private Sub CargarDatosVentaDia()
Dim RN As ADODB.Recordset
Dim Importe As Currency
Dim ColDespues As Collection
Dim Impor As Currency

    On Error GoTo eCargarDatosVentaDia
    Set RN = New ADODB.Recordset
    
    ListView1.ListItems.Clear
    DoEvents
    
    Cad = "select id ,inicial  ,codtrabaAper  FROM stpvdiacaja  where numtermi=" & vParamTPV.NumeroDeTerminal & " AND diacierre is null"
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    Text2(0).Tag = RN!ID
    Text1(0).Text = Format(RN!Inicial, FormatoImporte)
    'trabajador apertura no es el conectado
    Label1(9).Caption = ""
    If RN!codtrabaAper <> Val(Text2(1).Tag) Then
        Cad = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", CStr(RN!codtrabaAper))
        If Cad = "" Then Cad = "Error: " & RN!codtrabaAper
        Label1(9).Caption = Cad
    End If
    RN.Close
    
    lblIndic.Caption = "Leyendo en facturas ...."
    lblIndic.Refresh
    Cad = "select tipforpa,sum(totalfac) total from scafac,scafac1,sforpa" ',stipom "
    Cad = Cad & " WHERE scafac.codtipom = scafac1.codtipom and scafac.numfactu = scafac1.numfactu"
    Cad = Cad & " and scafac.fecfactu = scafac1.fecfactu AND scafac.codforpa=sforpa.codforpa "
    'Cad = Cad & " and scafac.codtipom=stipom.codtipom  "
    Cad = Cad & " and numtermi=" & vParamTPV.NumeroDeTerminal
    Cad = Cad & "  and fechaalb>=" & DBSet(Fecha, "F") & " group by 1 "
    '%=%=añadido 3,4
    Text1(5).Text = ""
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColDespues = New Collection
    If Not RN.EOF Then
        ListView1.ListItems.Add , , "Facturas"
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "   "
        ListView1.ListItems(ListView1.ListItems.Count).Bold = True

        Cad = ""
        
        While Not RN.EOF
            
            If RN!tipforpa = 0 Or RN!tipforpa = 6 Then
                Cad = IIf(RN!tipforpa = 0, "Efectivo", "Tarjeta")
                If RN!tipforpa = 0 Then Text1(5).Text = Format(RN!total, FormatoImporte)
                ListView1.ListItems.Add , , Cad
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = Format(RN!total, FormatoImporte)
            Else
                ColDespues.Add RN!tipforpa & "|" & Format(RN!total, FormatoImporte) & "|"
            End If
            RN.MoveNext
        Wend
    End If
    RN.Close
        
    lblIndic.Caption = "Tipo pago"
    lblIndic.Refresh
    If ColDespues.Count > 0 Then
        RN.Open "Select * from stippa", conn, adOpenKeyset, adLockPessimistic, adCmdText
        For NumRegElim = 1 To ColDespues.Count
            
            Cad = RecuperaValor(ColDespues.Item(NumRegElim), 1)
            Cad = " tipforpa =" & Cad
            RN.Find Cad, , adSearchForward, 0
            If RN.EOF Then
                Cad = "ERROR : " & RecuperaValor(ColDespues.Item(NumRegElim), 1)
            Else
                Cad = LCase(RN!destippa)
            End If
            ListView1.ListItems.Add , , Cad
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = RecuperaValor(ColDespues.Item(NumRegElim), 2)
            
            
        Next
        RN.Close
    End If
    'Por si ha creado albaranes
    lblIndic.Caption = "Leyendo albaranes"
    lblIndic.Refresh
    
''    'Antes 2022
''    Cad = " select codigiva,sum(importel) from slialb,sartic  where slialb.codartic=sartic.codartic and"
''    Cad = Cad & " (codtipom,numalbar) IN (select codtipom,numalbar from scaalb,sforpa where"
''    Cad = Cad & " scaalb.codforpa = sforpa.codforpa And tipforpa = 0"
''    Cad = Cad & " and fechaalb=" & DBSet(Fecha, "F") & " and esticket=1 and numtermi=" & vParamTPV.NumeroDeTerminal & ")"
''    '%=%=añadido el group by
''    Cad = Cad & " group by 1 order by 1 "
    
    'Ahora
    Cad = " select if(tipforpa=0,'Efectivo','--') forp, codigiva,sum(importel)"
    Cad = Cad & "  From scaalb, sforpa, slialb, sartic"
    Cad = Cad & " Where scaalb.codforpa = sforpa.codforpa AND slialb.codtipom=scaalb.codtipom AND slialb.numalbar=scaalb.numalbar"
    Cad = Cad & " AND slialb.codartic=sartic.codartic and fechaalb>=" & DBSet(Fecha, "F")
    Cad = Cad & " AND fechaalb <=" & DBSet(Now, "F")
    
    Cad = Cad & " AND esticket=1 and numtermi=" & vParamTPV.NumeroDeTerminal
    Cad = Cad & "  group by 1,2 order by 1 desc ,2"

    
    
    
    
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RN.EOF Then
        
        ListView1.ListItems.Add , , "Albaranes"
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "   "
        ListView1.ListItems(ListView1.ListItems.Count).Bold = True
        Impor = 0
        davidCodtipom = ""
        NumRegElim = 0
        While Not RN.EOF
                            
                If RN!forp <> davidCodtipom Then
                    If davidCodtipom <> "" Then ListView1.ListItems(NumRegElim).SubItems(1) = Format(Impor, FormatoImporte)
                    ListView1.ListItems.Add , , RN.Fields(0)
                    NumRegElim = ListView1.ListItems.Count
                    davidCodtipom = RN!forp
                    Impor = 0
                Else
                   
                End If
                Cad = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", RN.Fields(1))
                
                
                
                Importe = CCur(Cad) + 100
                Importe = Round((Importe * RN.Fields(2)) / 100, 2)
                Impor = Impor + Importe
                
            
                            
            RN.MoveNext
        Wend
        ListView1.ListItems(NumRegElim).SubItems(1) = Format(Impor, FormatoImporte)
    End If
    RN.Close
    davidCodtipom = ""
        
        
    'moviiemtnos de caja !!!
    'Dejarlo siempre para el final
    lblIndic.Caption = "leyendo movimientos caja...."
    lblIndic.Refresh
    Cad = "select entrada,sum(importe) total from stpventradassalidas where date(diahora)>=" & DBSet(Fecha, "F")
    Cad = Cad & " AND numtermi=" & vParamTPV.NumeroDeTerminal & " group by 1"
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    Importe = 0
    If Not RN.EOF Then
        ListView1.ListItems.Add , , "Movimientos"
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "   "
        ListView1.ListItems(ListView1.ListItems.Count).Bold = True
        While Not RN.EOF
            If RN!Entrada = 1 Then
                Cad = "Entrada"
                Impor = 1
            Else
                Cad = "Salida"
                Impor = -1
            End If
            ListView1.ListItems.Add , , Cad
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = Format(RN!total, FormatoImporte)
            Impor = Impor * RN!total
            Importe = Importe + Impor
            
            RN.MoveNext
        Wend
    End If
    RN.Close
    If Importe <> 0 Then Text1(2).Text = Format(Importe, FormatoImporte)

    
eCargarDatosVentaDia:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Cad & vbCrLf & Err.Description
        cmdAceptar.Enabled = False
    End If
    Set RN = Nothing
    Set ColDespues = Nothing
    lblIndic.Caption = ""
    ImporteFinal
End Sub



Private Sub CuadreVisible(visible As Boolean)
Dim B As Boolean
    
    If visible Then
        
        Cad = "select * from stpvtipomoneda ORDER BY valorcts" '  tipomoneda valorcts descripcion
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = -1
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            lblCuadre(NumRegElim).Caption = miRsAux!Descripcion
            lblCuadre(NumRegElim).Tag = miRsAux!tipomoneda
            txtCuadre(NumRegElim).Tag = miRsAux!valorcts
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
         If NumRegElim < Me.txtCuadre.Count - 1 Then B = True
            
    Else
        B = True
        NumRegElim = 0
    End If
        
    If B Then
        For davidNumalbar = NumRegElim To txtCuadre.Count - 1
            Me.txtCuadre(NumRegElim).visible = visible
            Me.txtCuadreImporte(NumRegElim).visible = visible
            Me.txtCuadre(NumRegElim).Text = ""
            Me.txtCuadreImporte(NumRegElim).Text = ""
            txtCuadreImporte(NumRegElim).Tag = 0
            Me.lblCuadre(NumRegElim).visible = visible
            
        Next
        

    End If
            
            
End Sub

Private Sub txtCuadre_GotFocus(Index As Integer)
    ConseguirFoco txtCuadre(Index), 3
End Sub

Private Sub txtCuadre_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCuadre_LostFocus(Index As Integer)
Dim ImporteTotal As Currency

    txtCuadre(Index).Text = Trim(txtCuadre(Index).Text)
    If txtCuadre(Index).Text <> "" Then
        If Not PonerFormatoEntero(txtCuadre(Index)) Then
            txtCuadre(Index).Text = ""
            txtCuadreImporte(Index).Text = ""
            txtCuadreImporte(Index).Tag = 0
            PonerFoco txtCuadre(Index)
        Else
            txtCuadreImporte(Index).Tag = (txtCuadre(Index).Text * txtCuadre(Index).Tag) / 100
            txtCuadreImporte(Index).Text = Format(txtCuadreImporte(Index).Tag, FormatoImporte)
        End If
    Else
        txtCuadreImporte(Index).Text = ""
        txtCuadreImporte(Index).Tag = 0
    End If
    
    ImporteTotal = 0
    For NumRegElim = 0 To Me.txtCuadre.Count - 1
        'txtCuadre(NumRegElim).Tag = miRsAux!valorcts
        Debug.Print txtCuadreImporte(Index).Tag
        If txtCuadreImporte(NumRegElim).Text <> "" Then ImporteTotal = ImporteTotal + txtCuadreImporte(NumRegElim).Tag
    Next
    If ImporteTotal = 0 Then
        Text1(1).Locked = False
        Text1(1).BackColor = vbWhite
    Else
        Text1(1).Locked = True
        Text1(1).BackColor = &HFFFF00
        Text1(1).Text = Format(ImporteTotal, FormatoImporte)
    End If
       
    ImporteFinal
    
End Sub


Private Sub ImporteFinal()
Dim Importe As Currency

    Importe = 0
    'Inicial
    If Text1(0).Text <> "" Then Importe = Importe + ImporteFormateado(Text1(0).Text)
    'Efectifo fra
    If Text1(5).Text <> "" Then Importe = Importe + ImporteFormateado(Text1(5).Text)
    'retirada
    If Text1(2).Text <> "" Then Importe = Importe + ImporteFormateado(Text1(2).Text)
    
    'SALDO caja
    Text1(4).Text = ""
    If Importe <> 0 Then Text1(4).Text = Format(Importe, FormatoImporte)
    
    'Diferencia
    Text1(3).Text = ""
    If Text1(1).Text <> "" Then
        Importe = ImporteFormateado(Text1(1).Text) - Importe
        If Importe <> 0 Then Text1(3).Text = Format(Importe, FormatoImporte)
        Text1(3).ForeColor = IIf(Importe < 0, vbRed, vbBlack)
    End If

    

End Sub



Private Sub Imprimir()
On Error GoTo eImprimir
    Screen.MousePointer = vbHourglass
    
    conn.Execute "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    conn.Execute "DELETE FROM tmpnlotes WHERE codusu =" & vUsu.Codigo
    conn.Execute "DELETE FROM tmpsliped WHERE codusu =" & vUsu.Codigo
    

    '
    'codusu fecha1 nombre1  nombre2 nombre3
    PonerTrabajadorConectado Cad
    Cad = vUsu.Codigo & "," & vParamTPV.NumeroDeTerminal & "," & DBSet(Text2(0).Text, "F") & "," & DBSet(Text2(1).Text, "T") & "," & DBSet(Text2(2).Text, "T") & "," & DBSet(Cad, "T")
    ' apert     efect   retira   saldo    caja      dif     INCIO sig
    'importe1 importe2 importe3 importe4 importe5 importeb1 importeb2
    Cad = Cad & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(5).Text, "N") & "," & DBSet(Text1(2).Text, "N")
    Cad = Cad & "," & DBSet(Text1(4).Text, "N") & "," & DBSet(Text1(1).Text, "N") & "," & DBSet(Text1(3).Text, "N")
    Cad = Cad & "," & IIf(Text1(6).Text = "", "null", DBSet(Text1(6).Text, "N"))
    Cad = Cad & "," & DBSet(Now, "FH")
    Cad = " VALUES (" & Cad & ")"
    
    
    Cad = "INSERT INTO tmpinformes (codusu ,codigo1,fecha1 ,nombre1  ,nombre2,nombre3 ,importe1 ,importe2, importe3 ,importe4 ,importe5 ,importeb1 ,importeb2 ,FECHA3) " & Cad
    conn.Execute Cad


    
    Cad = ""
    For NumRegElim = 0 To Me.txtCuadre.Count - 1
        If Me.txtCuadre(NumRegElim).visible Then
            If txtCuadre(NumRegElim).Text <> "" Then
                'lblCuadre(NumRegElim).Tag  ' = miRsAux!tipomoneda
                'tmpnlotes codusu numalbar fechaalb codprove numlinea nomartic codalmac cantidad
                Cad = Cad & ",  (" & vUsu.Codigo & ",'a','1972-04-12',1," & lblCuadre(NumRegElim).Tag & ","
                Cad = Cad & DBSet(lblCuadre(NumRegElim).Caption, "T") & "," & DBSet(txtCuadre(NumRegElim).Text, "N") & "," & DBSet(txtCuadreImporte(NumRegElim).Text, "N") & ")"
            End If
        End If
    Next
    If Cad <> "" Then
        Cad = "INSERT INTO tmpnlotes (codusu ,numalbar ,fechaalb ,codprove ,numlinea ,nomartic ,codalmac ,cantidad) VALUES " & Mid(Cad, 2)
        conn.Execute Cad
    End If
    
    '                          secuen    texto   negrit    importe
    'tmpsliped codusu numpedcl numlinea ampliaci codzona importel
    Cad = ""
    For NumRegElim = 1 To Me.ListView1.ListItems.Count
        Cad = Cad & ", (" & vUsu.Codigo & ",1," & NumRegElim & ","
        Cad = Cad & DBSet(Me.ListView1.ListItems(NumRegElim).Text, "T") & "," & IIf(ListView1.ListItems(NumRegElim).Bold, 1, 0)
        If ListView1.ListItems(NumRegElim).Bold Then
            Cad = Cad & ",0)"
        Else
            Cad = Cad & "," & DBSet(Me.ListView1.ListItems(NumRegElim).ListSubItems(1), "N") & ")"
        End If
    Next
    If Cad <> "" Then
        Cad = Mid(Cad, 2)
        Cad = "INSERT INTO tmpsliped (codusu ,numpedcl ,numlinea ,ampliaci ,codzona ,importel) VALUES " & Cad
        conn.Execute Cad
        
        
        Cad = "select " & vUsu.Codigo & ",1 entrada,@rownum:=@rownum+1 AS rownum ,descrip,3,if(entrada,importe,-importe)  from stpventradassalidas ,(SELECT @rownum:=" & NumRegElim & ") r"
        Cad = Cad & " where date(diahora)>=" & DBSet(Fecha, "F")
        Cad = Cad & " AND numtermi=" & vParamTPV.NumeroDeTerminal
        Cad = "INSERT INTO tmpsliped (codusu ,numpedcl ,numlinea ,ampliaci ,codzona ,importel)  " & Cad
        conn.Execute Cad
        
    End If


    'Si llega aqui, ya imprimimimos
    With frmImprimir
        .FormulaSeleccion = "{tmpInformes.codusu} = " & vUsu.Codigo
        .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
        .NumeroParametros = 1 'numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 3002
        .Titulo = "Cuadre caja"
        .NombreRPT = "rCierreCajaCuadre.rpt"
        .ConSubInforme = False
        .Show vbModal

    End With




eImprimir:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Screen.MousePointer = vbDefault
End Sub
