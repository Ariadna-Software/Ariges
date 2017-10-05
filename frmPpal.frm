VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPpal 
   BackColor       =   &H00858585&
   Caption         =   "Ariges 4"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   -990
   ClientWidth     =   15105
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageListB 
      Left            =   4920
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8688
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":909A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B8E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListPpal 
      Left            =   360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C2F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E418
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F4AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1053C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":13050
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":140E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":15174
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":16206
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":17298
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1832A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":193BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1A44E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1B4E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1C572
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1D604
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1E696
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1F728
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":210BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2791C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2BE1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2C830
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2FC22
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":36484
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3CCE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3DD78
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3EE0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3FE9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":40F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":41FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":48822
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4F084
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":558E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":56978
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":583FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":59E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5A196
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   30
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artículos"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Movimientos Art."
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proveedores"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ofertas Clientes"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Clientes"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaranes Clientes"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Cliente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas mostrador"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Proveedor"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaran Proveedor"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Factura Proveedor"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recepción Facturas Prov."
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimientos"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nº Serie"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Avisos"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gastos técnicos"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consulta precios / cliente"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Venta TPV"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar empresa"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agenda"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   8565
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpal.frx":609F8
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18574
            Text            =   "asdasd"
            TextSave        =   "asdasd"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "13:10"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   5640
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   52
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":63FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":65CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6BF6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6C97C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6D38E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6FB40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7041A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":70CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":715CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":71EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":728BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":72D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":72E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":72F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7304A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":73364
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":78F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":79998
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7A3AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7A4BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7AECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7B8E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7C2F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7C88C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7CBA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7CFF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7D44A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7D89C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7DCEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7E140
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7E592
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7E8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7EA06
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7ED20
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7F03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7F914
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":801EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":80508
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":80662
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8097C
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8138E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":81DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":827B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":831C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":83BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":845E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":84FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8B85C
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8D2DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8E370
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8ED82
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":96284
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListTPV 
      Left            =   360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9CAE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9E478
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9FE0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A179C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A312E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A4AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A6452
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A7DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AE646
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B3E38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListMAIL 
      Left            =   360
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B58BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B5D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B615E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B65B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B6A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B6E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B72A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B76F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B7B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":BDDE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":BE7F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C4A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CB2F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D83B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DEC18
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E547A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EBCDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EC12E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EC580
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EC9D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ECE24
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ED276
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ED6C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F32EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F413C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F4456
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F4770
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F4A8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnConfiguracion 
      Caption         =   "C&onfiguración"
      Begin VB.Menu mnConfParamGenerales 
         Caption         =   "Datos &Empresa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnConfParamAplic 
         Caption         =   "Parámetros &Aplicación"
      End
      Begin VB.Menu mnConTMovimiento 
         Caption         =   "Tipos &Movimiento"
      End
      Begin VB.Menu mnConfParamRpt 
         Caption         =   "Tipos de &Documentos"
      End
      Begin VB.Menu mnAridoc1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnAridoc1 
         Caption         =   "Configuración aridoc"
         Index           =   1
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnConfManteUsuarios 
         Caption         =   "Mantenimiento &Usuarios"
         HelpContextID   =   2
      End
      Begin VB.Menu mnNuevaEmpresa 
         Caption         =   "Creacion &nueva empresa"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnUsuarios 
         Caption         =   "Nuevo U&suario"
         Visible         =   0   'False
      End
      Begin VB.Menu mnPedirPwd 
         Caption         =   "Password requerido"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCambioEmpresa 
         Caption         =   "Cambiar Em&presa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnBarra17 
         Caption         =   "-"
      End
      Begin VB.Menu mnSeleccionarImpresora 
         Caption         =   "Seleccionar &Impresora"
      End
      Begin VB.Menu mnBarra12 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnAlmacen 
      Caption         =   "&Almacen"
      Begin VB.Menu mnDatosGenAlmacen 
         Caption         =   "&Datos Generales"
         Begin VB.Menu mnAlmMarcas 
            Caption         =   "&Marcas"
         End
         Begin VB.Menu mnAlmAlPropios 
            Caption         =   "Almacenes &Propios"
         End
         Begin VB.Menu mnAlmTipoUnidad 
            Caption         =   "Tipos &Unidad"
         End
         Begin VB.Menu mnTiposArticulos 
            Caption         =   "&Tipos Articulos"
         End
         Begin VB.Menu mnAlmUbicacion 
            Caption         =   "U&bicaciones"
         End
         Begin VB.Menu mnAlmFamiliaArticulo 
            Caption         =   "&Familias Artículos"
         End
         Begin VB.Menu mnAlmCategoria 
            Caption         =   "&Categorías"
         End
         Begin VB.Menu mnAlmArticulos 
            Caption         =   "&Artículos"
         End
         Begin VB.Menu mnAlmNumLotes 
            Caption         =   "&Numeros de lote"
         End
      End
      Begin VB.Menu mnAlmMovimientosAlm 
         Caption         =   "&Movimientos Almacen"
         Begin VB.Menu mnAlmTraspaso 
            Caption         =   "&Traspaso Almacenes"
         End
         Begin VB.Menu mnAlmTraspasoHco 
            Caption         =   "&Histórico Traspaso Almacenes"
         End
         Begin VB.Menu mnAlmMovimientos 
            Caption         =   "&Movimientos Almacen"
         End
         Begin VB.Menu mnAlmMovimientosHco 
            Caption         =   "H&istórico Movimientos Almacen"
         End
      End
      Begin VB.Menu mnAlmConsultas 
         Caption         =   "&Consultas"
         Begin VB.Menu mnAlmMovimArticulos 
            Caption         =   "Movimientos A&rticulos"
         End
         Begin VB.Menu mnAlmListMovim 
            Caption         =   "Listado &Movimientos"
         End
         Begin VB.Menu mnAlmMovimArticulosSt 
            Caption         =   "Movimientos stock desde inventario"
         End
         Begin VB.Menu mnAlmControlStockDesdeInv 
            Caption         =   "Listado control stock"
         End
         Begin VB.Menu mnAlmListInactivos 
            Caption         =   "Listado Articulos &Inactivos"
         End
         Begin VB.Menu mnAlmListComponentes 
            Caption         =   "Listado Articulos &Componentes"
         End
         Begin VB.Menu mnAlmListValoracion 
            Caption         =   "Listado Valoración &Stocks"
         End
         Begin VB.Menu mnAlmListMaxMin 
            Caption         =   "Inf. Stocks Máximos-Mínimos"
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Inf. Stocks a una &Fecha"
            Index           =   0
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Stocks por meses"
            Index           =   1
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Alertas punto pedido"
            Index           =   2
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Informe reposición almacen"
            Index           =   3
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Listado stock mínimo"
            Index           =   4
         End
      End
      Begin VB.Menu mnAlmInventario 
         Caption         =   "&Inventario"
         Begin VB.Menu mnAlmTomaInven 
            Caption         =   "&Toma de inventario"
         End
         Begin VB.Menu mnAlmEntradaInve 
            Caption         =   "&Entrada existencia real"
         End
         Begin VB.Menu mnAlmListadoInve 
            Caption         =   "&Listado diferencias"
         End
         Begin VB.Menu mnAlmActualizarInve 
            Caption         =   "Actualizar &diferencias"
         End
         Begin VB.Menu mnAlmValoracionInve 
            Caption         =   "&Valoración stocks inventariados"
         End
         Begin VB.Menu mnRectifInve 
            Caption         =   "Rectificar último inventario"
            Index           =   0
         End
         Begin VB.Menu mnRectifInve 
            Caption         =   "Inventariar artículo"
            Index           =   1
         End
         Begin VB.Menu mnRecalPrSt 
            Caption         =   "Recálculo precio estándar"
            Index           =   0
         End
         Begin VB.Menu mnRecalPrSt 
            Caption         =   "Recálculo precio medio ponderado"
            Index           =   1
         End
         Begin VB.Menu mnRecalPrSt 
            Caption         =   "Recálculo ultimo precio compra"
            Index           =   2
         End
         Begin VB.Menu mnBarra2 
            Caption         =   "-"
         End
         Begin VB.Menu mnAlmHcoInven 
            Caption         =   "&Histórico inventario"
         End
      End
      Begin VB.Menu mnTelematel 
         Caption         =   "Telematel"
         Index           =   0
      End
      Begin VB.Menu mnTelematel 
         Caption         =   "Comunicación datos grupo"
         Index           =   1
      End
   End
   Begin VB.Menu mnFacturacion 
      Caption         =   "&Facturación"
      Begin VB.Menu mnFacDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnFacActividades 
            Caption         =   "Activi&dades"
         End
         Begin VB.Menu mnFacZonas 
            Caption         =   "&Zonas"
         End
         Begin VB.Menu mnFacRutas 
            Caption         =   "&Rutas"
         End
         Begin VB.Menu mnPortes 
            Caption         =   "Portes"
         End
         Begin VB.Menu mnDtoCantidad 
            Caption         =   "Descuento por cantidad"
         End
         Begin VB.Menu mnFacFormasEnvio 
            Caption         =   "Formas de &Envio"
         End
         Begin VB.Menu mnFacFormasPago 
            Caption         =   "Formas de &Pago"
         End
         Begin VB.Menu mnFacBancosPropios 
            Caption         =   "&Bancos Propios"
         End
         Begin VB.Menu mnFacSituaciones 
            Caption         =   "&Situaciones Especiales"
         End
         Begin VB.Menu mnFacAgentesCom 
            Caption         =   "Agentes &Comerciales"
         End
         Begin VB.Menu mnFacClientesV1 
            Caption         =   "Clientes &Varios"
         End
         Begin VB.Menu mnFacClientes 
            Caption         =   "Cl&ientes"
         End
         Begin VB.Menu mnFacClientesPot 
            Caption         =   "Clientes potenciales"
         End
         Begin VB.Menu mnFacCartas 
            Caption         =   "Tipos de C&artas"
         End
         Begin VB.Menu mnFacIncidencias 
            Caption         =   "&Incidencias"
         End
      End
      Begin VB.Menu mnFacInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Clientes Inacti&vos"
            Index           =   0
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "&Clientes"
            Index           =   1
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "&Altas Clientes"
            Index           =   2
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "&Etiquetas de clientes"
            Index           =   3
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Car&tas a clientes"
            Index           =   4
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "&Etiquetas de bultos"
            Index           =   5
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Listado teléfonos x cliente"
            Index           =   6
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Listado cuotas telefonía"
            Index           =   7
         End
      End
      Begin VB.Menu mnFacPreciosDtos 
         Caption         =   "&Precios y Descuentos"
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Tarifas Venta"
            Index           =   0
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Lista Precios"
            Index           =   1
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Precios &Especiales"
            Index           =   2
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Promociones"
            Index           =   3
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Descuentos Familia/Marca"
            Index           =   4
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Descuento por actividad"
            Index           =   5
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Bonificaciones Factura"
            Index           =   6
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Actualizar precios"
            Index           =   8
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Copiar precios desde compra"
            Index           =   9
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Control margenes tarifas"
            Index           =   11
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Corrección errores y actualización tarifas"
            Index           =   12
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Control error descuentos por cliente"
            Index           =   13
         End
      End
      Begin VB.Menu mnFacOfert 
         Caption         =   "&Ofertas"
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Mantenimiento Ofertas"
            Index           =   0
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Grupo de Plantillas"
            Index           =   1
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "Entrada de  &Plantillas"
            Index           =   2
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "Ofertas E&fectuadas"
            Index           =   3
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Histórico  Ofertas"
            Index           =   5
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Traspaso a Histórico"
            Index           =   6
         End
      End
      Begin VB.Menu mnFacPed 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Mantenimiento Pedidos"
            Index           =   0
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Histórico Pedidos Anulados"
            Index           =   1
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Cartas Confirmacion de Pedidos"
            Index           =   3
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe &Pedidos por Articulo"
            Index           =   4
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe P&edidos por Cliente"
            Index           =   5
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe &Disponibilidad Stocks"
            Index           =   6
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Impresion pedidos por zona"
            Index           =   7
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Consulta precios / cliente"
            Index           =   9
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Estadísticas consultas precios/cliente"
            Index           =   10
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Devolución material"
            Index           =   12
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Pedido presupuestos"
            Index           =   13
         End
      End
      Begin VB.Menu mnFacAlbaran 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnFacEntAlbaran 
            Caption         =   "&Mantenimiento Albaranes"
         End
         Begin VB.Menu mnFacAlbDevolucion 
            Caption         =   "Albaranes de devolución"
         End
         Begin VB.Menu mnAlbaranesB 
            Caption         =   "Albaranes presupuestos *"
         End
         Begin VB.Menu mnFacAlbxArtic 
            Caption         =   "Informe &Albaranes por Articulo"
         End
         Begin VB.Menu mnFacIncumPlazos 
            Caption         =   "Inf. Incumplimiento Plazos &Ent."
         End
         Begin VB.Menu mnFacHcoAlbaranes 
            Caption         =   "&Histórico Albaranes Anulados"
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Situación albaranes"
            Index           =   0
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Control de albaranes"
            Index           =   1
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Control albaranes facturados"
            Index           =   2
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Impresion albaranes transporte"
            Index           =   3
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Control direcciones de envio"
            Index           =   4
         End
         Begin VB.Menu mnBarra5 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacPreFacturar 
            Caption         =   "&Previsión Facturación"
         End
         Begin VB.Menu mnFacFacturarAlb 
            Caption         =   "&Facturación de Albaranes"
         End
         Begin VB.Menu mnFacturarCliente 
            Caption         =   "Facturar cliente"
         End
         Begin VB.Menu mnFacAlbMostrador 
            Caption         =   "Facturas de Mo&strador"
         End
         Begin VB.Menu mnFacturarPresupuestos 
            Caption         =   "Facturar presupuestos *"
         End
         Begin VB.Menu mnFacAlbRectifica 
            Caption         =   "Facturas &Rectificativas"
         End
         Begin VB.Menu mnFacHcoFacturas 
            Caption         =   "His&tórico Albaran/Factura"
         End
         Begin VB.Menu mnFacReImpFactu 
            Caption         =   "Re&imprimir Facturas"
         End
         Begin VB.Menu mnEnvioFactuasMail 
            Caption         =   "Enviar facturas por e&mail"
            Index           =   0
         End
         Begin VB.Menu mnEnvioFactuasMail 
            Caption         =   "Facturacion web/electrónica"
            Index           =   1
         End
         Begin VB.Menu mnServicios 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Albaranes de servicio"
            Index           =   1
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Facturación de servicios"
            Index           =   2
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Albaranes internos"
            Index           =   3
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Facturacion albaranes  internos"
            Index           =   4
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Listado albaranes internos"
            Index           =   5
         End
         Begin VB.Menu mnTicket 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnTicket 
            Caption         =   "Contabilizar facturas tickets agrupados"
            Index           =   1
         End
         Begin VB.Menu mnTicket 
            Caption         =   "Listado tickets facturados"
            Index           =   2
         End
         Begin VB.Menu mnBarra9 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
      End
      Begin VB.Menu mnTelefonia 
         Caption         =   "&Telefonía"
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Albaranes de telefonía"
            Index           =   0
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Importar fichero"
            Index           =   2
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Datos pendientes facturar"
            Index           =   3
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Consumos"
            Index           =   5
            Begin VB.Menu mnTelefonia3 
               Caption         =   "Conceptos"
               Index           =   0
            End
            Begin VB.Menu mnTelefonia3 
               Caption         =   "Descuentos"
               Index           =   1
            End
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Cuotas"
            Index           =   6
            Begin VB.Menu mnTelefonia4 
               Caption         =   "Conceptos"
               Index           =   0
            End
            Begin VB.Menu mnTelefonia4 
               Caption         =   "Descuentos"
               Index           =   1
            End
            Begin VB.Menu mnTelefonia4 
               Caption         =   "Cuotas propias cooperativa"
               Index           =   2
            End
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Cargos varios"
            Index           =   7
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Modificación masiva cuotas"
            Index           =   8
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Comparativa descuentos"
            Index           =   10
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Facturacion x soporte"
            Index           =   11
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Resumen por soporte"
            Index           =   12
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Datos importación fichero"
            Index           =   14
         End
      End
      Begin VB.Menu mnAgua 
         Caption         =   "Agua"
         Begin VB.Menu mnAguaLin 
            Caption         =   "Contadores"
            Index           =   0
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Importar fichero"
            Index           =   1
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Facturar"
            Index           =   2
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Resumen facturación"
            Index           =   3
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Listado facturación por periodo"
            Index           =   4
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Listado contadores exportación"
            Index           =   5
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Modificar cuota Varios"
            Index           =   6
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Declaración detallada ejercicio"
            Index           =   7
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Calibres"
            Index           =   9
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Parámetros"
            Index           =   10
         End
      End
      Begin VB.Menu mnTratamientosRaiz 
         Caption         =   "Tratamientos"
         Begin VB.Menu mnTratamientos 
            Caption         =   "Mto materias activas"
            Index           =   0
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Mantenimiento ADR"
            Index           =   1
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Plagas"
            Index           =   2
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Flotas"
            Index           =   3
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Tratamientos"
            Index           =   4
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Partes trabajo"
            Index           =   5
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Listado fitosanitarios/campos"
            Index           =   6
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Vacio y NO visible"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Ajuste compras tratamientos"
            Index           =   9
         End
      End
      Begin VB.Menu mnObra1 
         Caption         =   "Obras"
         Begin VB.Menu mnobra 
            Caption         =   "Capítulos"
            Index           =   0
         End
         Begin VB.Menu mnobra 
            Caption         =   "Actuaciones"
            Index           =   1
         End
         Begin VB.Menu mnobra 
            Caption         =   "Partes de trabajo"
            Index           =   2
         End
         Begin VB.Menu mnobra 
            Caption         =   "Mto tipos órdenes de trabajo"
            Index           =   3
         End
         Begin VB.Menu mnobra 
            Caption         =   "Reloj"
            Index           =   4
         End
         Begin VB.Menu mnobra 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnobra 
            Caption         =   "Listado compras-ventas actuacion"
            Index           =   6
         End
         Begin VB.Menu mnobra 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnobra 
            Caption         =   "Imprimir certificación"
            Index           =   8
         End
      End
      Begin VB.Menu mnHuertos 
         Caption         =   "Gestion parcelas"
         Begin VB.Menu mnHuertos1 
            Caption         =   "Listado campos-hanegadas"
            Index           =   0
         End
         Begin VB.Menu mnHuertos1 
            Caption         =   "Facturación derrama"
            Index           =   1
         End
      End
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnFacEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnFacEstVentaCliente 
            Caption         =   "&Ventas por cliente"
         End
         Begin VB.Menu mnFacEstVentaTraba 
            Caption         =   "Ventas por &trabajador"
         End
         Begin VB.Menu mnFacEstVentaMes 
            Caption         =   "Ventas por &meses"
         End
         Begin VB.Menu mnFacEstVentaFam 
            Caption         =   "Ventas por &familia  /  Artículo"
         End
         Begin VB.Menu mnEstVtasArituclo 
            Caption         =   "Ventas por artículo"
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por proveedor"
            Index           =   0
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por &agente"
            Index           =   1
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "&Detalle facturación"
            Index           =   2
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Mar&gen ventas por artículo "
            Index           =   3
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por tipo de precio"
            Index           =   4
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Articulos más vendidos"
            Index           =   5
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas familia agrupado"
            Index           =   6
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por tipo de pedido"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnCompras 
      Caption         =   "&Compras"
      Begin VB.Menu mnComDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnComProveedores 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnComProveVarios 
            Caption         =   "Proveedores &Varios"
         End
         Begin VB.Menu mnComDirecciones 
            Caption         =   "&Direcciones"
         End
      End
      Begin VB.Menu mnComInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnComInfVarios1 
            Caption         =   "&Proveedores"
            Index           =   0
         End
         Begin VB.Menu mnComInfVarios1 
            Caption         =   "&Etiquetas de proveedores"
            Index           =   1
         End
         Begin VB.Menu mnComInfVarios1 
            Caption         =   "&Cartas a Proveedores"
            Index           =   2
         End
         Begin VB.Menu mnComInfVarios1 
            Caption         =   "Etiquetas de bultos"
            Index           =   3
         End
      End
      Begin VB.Menu mnComPreciosDtos 
         Caption         =   "Precios y &Descuentos"
         Begin VB.Menu mnComPreProv 
            Caption         =   "P&recios Proveedor"
            Index           =   0
         End
         Begin VB.Menu mnComPreProv 
            Caption         =   "Descuentos Pro&veedor"
            Index           =   1
         End
         Begin VB.Menu mnComPreProv 
            Caption         =   "Copiar precios desde venta"
            Index           =   2
         End
         Begin VB.Menu mnComPreProv 
            Caption         =   "Actualizar precios"
            Index           =   3
         End
      End
      Begin VB.Menu mnComPedidos 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "Mant. &Pedidos Proveedor"
            Index           =   0
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "&Histórico Pedidos Anulados"
            Index           =   1
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "List. &Material pendiente de recibir"
            Index           =   2
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "Propuesta pedido"
            Index           =   4
         End
      End
      Begin VB.Menu mnComAlbaranes 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnComAlbMan 
            Caption         =   "&Mant. Albaranes Proveedor"
         End
         Begin VB.Menu mnComHcoAlbaranes 
            Caption         =   "&Histórico Albaranes Anulados"
         End
         Begin VB.Menu mnComPteFacturar 
            Caption         =   "List. &Pendiente de facturar"
         End
         Begin VB.Menu mnBarra7 
            Caption         =   "-"
         End
         Begin VB.Menu mnComFacturar 
            Caption         =   "&Recepción Facturas"
         End
         Begin VB.Menu mnComHcoFacturas 
            Caption         =   "&Histórico Albaran/Factura"
         End
         Begin VB.Menu mnBarra15 
            Caption         =   "-"
         End
         Begin VB.Menu mnComContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
         Begin VB.Menu mnComCtrlAlb 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnComCtrlAlb 
            Caption         =   "Control albaranes"
            Index           =   1
         End
         Begin VB.Menu mnComCtrlAlb 
            Caption         =   "Control albaranes facturados"
            Index           =   2
         End
      End
      Begin VB.Menu mnProcesoLiquidacionProveedores 
         Caption         =   "Liquidación proveedores"
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Cambiar precios"
            Index           =   0
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Liquidacion proveedores"
            Index           =   1
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Impresion facturas"
            Index           =   2
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Asociar albaranes compras / ventas"
            Index           =   3
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Listado asociaciones albaranes"
            Index           =   4
         End
      End
      Begin VB.Menu Barra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnComEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras por &Proveedor"
            Index           =   0
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras por &Familia/Artíc."
            Index           =   1
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras por &meses"
            Index           =   2
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "&Albaranes por Proveedor"
            Index           =   3
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Informe previsión pagos"
            Index           =   4
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras Proveedor-Marca-Familia"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnAdministracion 
      Caption         =   "A&dministración"
      Begin VB.Menu mnAdmDatosGen 
         Caption         =   "&Datos Generales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnAdmTrabajadores 
         Caption         =   "&Trabajadores"
      End
      Begin VB.Menu mnAdmGastosTec 
         Caption         =   "&Gastos Técnicos"
      End
      Begin VB.Menu mnAdmNominas 
         Caption         =   "&Nominas y Gastos"
      End
      Begin VB.Menu mnInformesAdm 
         Caption         =   "Informes"
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Beneficio por proveedor"
            Index           =   3
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Beneficio por cliente"
            Index           =   4
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Beneficio marca-agente-proveedor"
            Index           =   5
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Informe de artículos en promocion"
            Index           =   6
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Informe ventas articulos con dto. especial"
            Index           =   7
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Ventas trabajador / Dia"
            Index           =   8
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Ventas por forma de pago"
            Index           =   9
         End
      End
      Begin VB.Menu mnAdmAgentes 
         Caption         =   "Agentes"
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Resumen ventas - agente"
            Index           =   0
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Beneficio por agente"
            Index           =   1
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Ventas agente - trabajador"
            Index           =   2
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Pagos comisiones"
            Index           =   4
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Generar pagos comisiones"
            Index           =   5
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Listado comisiones ECO"
            Index           =   7
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Listado agente-familia-marca"
            Index           =   8
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Listado agente-marca-familia"
            Index           =   9
         End
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Calculo de &riesgo"
         Index           =   1
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Informe ventas a credito"
         Index           =   2
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Informe de previsión tesoreria"
         Index           =   3
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Corrección costes varios factura"
         Index           =   4
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Modificar coste estadistica ventas"
         Index           =   5
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Gestión de flotas-Maquinaria"
         Index           =   6
         Begin VB.Menu mnFlotas 
            Caption         =   "Registro "
            Index           =   0
         End
         Begin VB.Menu mnFlotas 
            Caption         =   "Mantenimiento de flotas"
            Index           =   1
         End
         Begin VB.Menu mnFlotas 
            Caption         =   "Mantenimiento de conceptos"
            Index           =   2
         End
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Comunicación datos seguro"
         Index           =   7
      End
   End
   Begin VB.Menu mnMantenimientos 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnManTiposContrato 
         Caption         =   "&Tipos de Contrato"
      End
      Begin VB.Menu mnManEntrada 
         Caption         =   "&Entrada Mantenimientos"
      End
      Begin VB.Menu mnBarra8 
         Caption         =   "-"
      End
      Begin VB.Menu mnManListado 
         Caption         =   "&Listado Mantenimientos"
      End
      Begin VB.Menu mnManRevisiones 
         Caption         =   "Listado &Revisiones Mant."
      End
      Begin VB.Menu mnManFichas 
         Caption         =   "&Fichas Mantenimientos"
      End
      Begin VB.Menu mnManAltas 
         Caption         =   "List. &Altas Mantenimientos"
      End
      Begin VB.Menu mnInfTeoMant 
         Caption         =   "Informe teórico mantenimientos"
      End
      Begin VB.Menu mnEtiqMante 
         Caption         =   "Etiquetas de mantenimientos"
      End
      Begin VB.Menu mnBarra30 
         Caption         =   "-"
      End
      Begin VB.Menu mnCartaRenovaMante 
         Caption         =   "Carta renovación"
      End
      Begin VB.Menu mnTraspasoMante 
         Caption         =   "Traspaso siguiente a actual"
      End
      Begin VB.Menu mnBarra32 
         Caption         =   "-"
      End
      Begin VB.Menu mnHcoMaten 
         Caption         =   "Histórico mantenimientos anulados"
      End
      Begin VB.Menu mnInfManteAnulados 
         Caption         =   "Informe mantenimientos anulados"
      End
      Begin VB.Menu mnBarra13 
         Caption         =   "-"
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "&Previsión Facturación"
         Index           =   0
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "Fac&turación  Mantenimientos"
         Index           =   1
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "Prevision facturacion Renting y servicios"
         Index           =   3
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "Facturacion renting y servicios"
         Index           =   4
      End
   End
   Begin VB.Menu mnReparaciones 
      Caption         =   "&Reparaciones"
      Begin VB.Menu mnRepEntReparacion 
         Caption         =   "&Mant.  Reparaciones"
      End
      Begin VB.Menu mnRepControlRep 
         Caption         =   "C&ontrol Reparaciones"
      End
      Begin VB.Menu mnRepNumSerie 
         Caption         =   "Mant. &Nº Serie"
      End
      Begin VB.Menu mnRepMotivosBaja 
         Caption         =   "Motivos &baja equipos"
      End
      Begin VB.Menu mnRepMotivosPend 
         Caption         =   "Motivos &Pend. Rep."
      End
      Begin VB.Menu mnRepHistorico 
         Caption         =   "&Histórico de Reparaciones"
      End
      Begin VB.Menu mnManServicioAsisTecn 
         Caption         =   "Servicios asistencia técnica"
      End
      Begin VB.Menu mnTiposAveria 
         Caption         =   "Tipos averia"
      End
      Begin VB.Menu mnTrabaRealiz 
         Caption         =   "Trabajos realizados"
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "Albarán orden de trabajo"
         Index           =   1
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "Albarán trabajo exterior"
         Index           =   2
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "Albarán de reparación"
         Index           =   3
      End
      Begin VB.Menu Barra9 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepListRepxDia 
         Caption         =   "Listado Rep. del &Dia"
      End
      Begin VB.Menu mnRepListRepxClien 
         Caption         =   "Listado Rep. por &Cliente"
      End
      Begin VB.Menu mnRepListFrecuen 
         Caption         =   "F&recuencia de reparaciones"
      End
      Begin VB.Menu mnEstadisticaReparacionTecnico 
         Caption         =   "Estadística reparaciones técnico"
      End
      Begin VB.Menu mnListadoReparacionesEfectuadas 
         Caption         =   "Listado reparaciones efectuadas"
      End
      Begin VB.Menu mnRepGarantprove 
         Caption         =   "Reparaciones garantia proveedor"
      End
      Begin VB.Menu Barra14 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepAlbaranes 
         Caption         =   "Mant. &Albaranes Rep."
      End
      Begin VB.Menu mnRepPrevFact 
         Caption         =   "Pre&visión Facturación"
      End
      Begin VB.Menu mnRepFactAlb 
         Caption         =   "&Facturación Reparaciones"
      End
      Begin VB.Menu mnBarra14 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepAvisos 
         Caption         =   "Av&isos de clientes"
      End
      Begin VB.Menu mnRepListAvisosPtes 
         Caption         =   "&Listado de avisos pendientes"
      End
      Begin VB.Menu mnBorrarAvisosCerrados 
         Caption         =   "Borre avisos cerrados"
      End
      Begin VB.Menu mnbarra33 
         Caption         =   "-"
      End
      Begin VB.Menu mnFrecuencias 
         Caption         =   "Frecuencias"
      End
   End
   Begin VB.Menu mnCRMmenu 
      Caption         =   "CRM"
      Begin VB.Menu mnCRM 
         Caption         =   "Mantenimiento acciones comerciales"
         Index           =   0
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Tipos acciones comerciales"
         Index           =   1
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Generar acciones comerciales"
         Index           =   2
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Impresion masiva"
         Index           =   3
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Impresion resumen CRM"
         Index           =   4
      End
   End
   Begin VB.Menu mnproduccion 
      Caption         =   "Producción"
      Begin VB.Menu mnproduccion1 
         Caption         =   "Órdenes producción"
         Index           =   0
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Ordenes de envasado"
         Index           =   1
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Descripción costes tasas"
         Index           =   2
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Registro trazabilidad"
         Index           =   3
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Parámetros cálidad"
         Index           =   4
      End
   End
   Begin VB.Menu mnTPV 
      Caption         =   "&Punto de Venta"
      Begin VB.Menu mnTPVLinea 
         Caption         =   "Pantalla de &venta"
         Index           =   0
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "&Cierre de caja"
         Index           =   1
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "Etiquetas estantería"
         Index           =   2
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "&Parámetros generales TPV"
         Index           =   4
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "Parámetros &terminales TPV"
         Index           =   5
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnAgenda 
         Caption         =   "&Agenda"
      End
      Begin VB.Menu mnVerAvisos 
         Caption         =   "A&visos"
      End
      Begin VB.Menu mnLlamadas 
         Caption         =   "Llamadas"
         Index           =   0
      End
      Begin VB.Menu mnLlamadas 
         Caption         =   "Concepto llamadas"
         Index           =   1
      End
      Begin VB.Menu mnBackUp 
         Caption         =   "&Copia Seguridad local"
      End
      Begin VB.Menu mnEliminarFacturas 
         Caption         =   "&Borre Facturas y Movimientos"
      End
      Begin VB.Menu mnRevisarMultibase 
         Caption         =   "Revisar caracteres especiales"
      End
      Begin VB.Menu mnManteneLOG 
         Caption         =   "Acciones realizadas"
      End
      Begin VB.Menu mnAridocFacturas 
         Caption         =   "Traspaso Aridoc"
      End
      Begin VB.Menu mnUtiDeclaraLOM 
         Caption         =   "Lotes fitosanitarios subvencionados"
         Index           =   0
      End
      Begin VB.Menu mnUtiDeclaraLOM 
         Caption         =   "Declaración ROPO"
         Index           =   1
      End
      Begin VB.Menu mnArticulos 
         Caption         =   "Acciones artículos"
         Begin VB.Menu mnArticulos2 
            Caption         =   "Eliminar articulos"
            Index           =   0
         End
         Begin VB.Menu mnArticulos2 
            Caption         =   "Cambiar familia / marca / proveedor"
            Index           =   1
         End
         Begin VB.Menu mnArticulos2 
            Caption         =   "Cambiar codigo articulo-referencia"
            Index           =   2
         End
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Listado albaranes-pedidos anulados"
         Index           =   0
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Comprobar cuenta bancaria/NIF"
         Index           =   1
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Traspaso contados ruta"
         Index           =   2
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Eliminar presupuestos"
         Index           =   3
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Configurar PDFs ver articulo"
         Index           =   4
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Exportar albaranes de servicio"
         Index           =   5
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Exportar email csv"
         Index           =   6
      End
      Begin VB.Menu mnBarra19 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiBuscar 
         Caption         =   "&Buscar..."
         Begin VB.Menu mnUtiBuscarErrFac 
            Caption         =   "&Errores en Nº Factura clientes"
         End
         Begin VB.Menu mnUtiBuscarPteCon 
            Caption         =   "Facturas pendientes de &contabilizar"
            Begin VB.Menu mnUtiBuscarErrConCli 
               Caption         =   "&Clientes"
            End
            Begin VB.Menu mnUtiBuscarErrConPro 
               Caption         =   "&Proveedores"
            End
         End
      End
      Begin VB.Menu mnCambioPwd 
         Caption         =   "Cambiar contraseña"
      End
      Begin VB.Menu mnBarra20 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiUsuActivos 
         Caption         =   "&Usuarios activos"
      End
      Begin VB.Menu mnUtiConnActivas 
         Caption         =   "&Conexiones activas"
      End
      Begin VB.Menu mnBarra21 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnUtiMensInt 
         Caption         =   "&Mensajeria interna"
         Visible         =   0   'False
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "&Nuevo"
            Index           =   0
         End
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "&Enviar/Recibir"
            Index           =   1
         End
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "&Tipo de mensaje"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnSoporte2 
      Caption         =   "&Soporte"
      Begin VB.Menu mnSoporte 
         Caption         =   "Ayuda"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Enviar Mail"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Web Ariadna Software"
         Index           =   4
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Comprobar version operativa"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Acerca de ..."
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Dim PrimeraVez As Boolean

Dim TieneEditorDeMenus As Boolean



Private Sub SituarArriba()
    On Error Resume Next
    Me.Top = 0
    Me.Left = 0
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub MDIForm_Activate()
Dim b As Boolean
 

   ' AvisosPendientes = False
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
       ' AvisosPendientes = TieneAvisosPendientes()
        If vParamAplic.NumeroInstalacion = 4 Then SituarArriba
        
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not vParam Is Nothing Then
        If vParam.Modificado Then
          'Poner datos visible del form
           
           vParam.Modificado = False
        End If
    End If
    
    PonerDatosVisiblesForm
    
    '-- Control de si se utilizan servicios o no ( si es que no no se muestra el menú) hemos fichado gente nueva para la copa
    '   el situarlo aqui hace que no haya que salir y entrar en el programa si se
    b = DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "ALI", "T") <> ""
    PuntoDeMenuVisible mnServicios(3), b
    PuntoDeMenuVisible mnServicios(4), b And vParamAplic.NumeroInstalacion <> 4
    PuntoDeMenuVisible mnServicios(5), b
    PuntoDeMenuVisible mnServicios(1), vParamAplic.Servicios
    PuntoDeMenuVisible mnServicios(2), vParamAplic.Servicios And vParamAplic.NumeroInstalacion <> 4  'EULER NO los factura
    b = b Or vParamAplic.Servicios
    PuntoDeMenuVisible mnServicios(0), b  'la barra
    
    
    'vParamAplic.Reparaciones
    
    'MAntenimientos y reparaciones
    'mnMantenimientos.visible = vParamAplic.Mantenimientos
    'mnReparaciones.visible = vParamAplic.Reparaciones
    PuntoDeMenuVisible mnReparaciones, vParamAplic.Reparaciones
    PuntoDeMenuVisible mnMantenimientos, vParamAplic.Mantenimientos
    
    
    '-- Eliminamos frecuencias de momento
    'mnFrecuencias.visible = vParamAplic.Frecuencias
    PuntoDeMenuVisible mnFrecuencias, vParamAplic.Frecuencias
    Me.mnbarra33.visible = mnFrecuencias.visible
    
    '--------------------
    'TElefonia
    mnTelefonia.visible = vParamAplic.TieneTelefonia2 > 0
    If vParamAplic.TieneTelefonia2 > 0 Then
        'Catadau
        'If vParamAplic.TieneTelefonia2 = 2 Then Me.mnTelefonia1(2).Caption = "Importar fichero COARVAL"
'
'        If vParamAplic.TieneTelefonia2 = 1 Then
'            Me.mnTelefonia1(5).Caption = "Descuentos consumo"
'            Me.mnTelefonia1(6).Caption = "Descuentos cuotas"
'        Else
'            Me.mnTelefonia1(5).Caption = "Conceptos consumo"
'            Me.mnTelefonia1(6).Caption = "Conceptos cuotas"
'        End If
    End If

    '-- Contabilizacion tickets agrupados
    mnTicket(0).visible = vParamAplic.ContabilizarTicketAgrupados
    mnTicket(1).visible = vParamAplic.ContabilizarTicketAgrupados
    mnTicket(2).visible = vParamAplic.ContabilizarTicketAgrupados
    
        
    'Los albaranes y facturas en "B"
    'seran visibles si esta creado el tipo movimiento y tene contabilidad B
    b = DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "ALZ", "T") <> ""
    b = b And vParamAplic.ContabilidadB > 0
    PuntoDeMenuVisible Me.mnAlbaranesB, b
    PuntoDeMenuVisible mnFacturarPresupuestos, b
   

       
    'De momento:
    PuntoDeMenuVisible Me.mnAridoc1(0), True
    PuntoDeMenuVisible Me.mnAridoc1(1), True
    
       
    'Produccion
    PuntoDeMenuVisible Me.mnproduccion, vParamAplic.Produccion
       
       
    PuntoDeMenuVisible Me.mnCRMmenu, vParamAplic.TieneCRM
       
    'Flotas
    
    PuntoDeMenuVisible Me.mnAdministra(6), vParamAplic.GestionFlotas
       
    'Obras
    Me.mnObra1.visible = vParamAplic.HayDeparNuevo = 2
    'Mensajeria
    mnUtiMensInt.visible = False
    mnBarra21.visible = False
    
 '   If AvisosPendientes Then
 '       If MsgBox("Tiene avisos pendientes. ¿Quiere verlos ahora?", vbQuestion + vbYesNo) = vbYes Then
 '           'Mostrare la pantalla de avisos pendientes
 '           frmAlertas.Show vbModal
 '       End If
 '   End If
    '-- Descriptores especiales (Vrs 4.0.9)
    If vParamAplic.Descriptores Then
        mnAlmTipoUnidad.Caption = "Formatos"
        mnTiposArticulos.Caption = "Modelos"
        mnAlmFamiliaArticulo.Caption = "Categorias Art."
        mnAlmCategoria.visible = False
    End If
    
    'Para este usuario y esta empresa unos avlores al usuario
    vUsu.FijarOtrosValoresUsuario
    
    'mnAlmMovimArticulosSt.visible = False
    'Dim TieneDevolucionRMA As Boolean  'Si tiene el tipo de movimiento PEW''
    b = False
    If DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "PEW", "T") <> "" Then b = True
    PuntoDeMenuVisible Me.mnFacPedidos(11), b
    PuntoDeMenuVisible Me.mnFacPedidos(12), b
    b = False
    'If vParamAplic.AlmacenB >= 0 Then
    If vParamAplic.NumeroInstalacion = 2 Then
        'Tiene almacen B
        If DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "PEZ", "T") <> "" Then
            If vUsu.AlmacenPorDefecto2 = CStr(vParamAplic.AlmacenB) Then b = True
        End If
    End If
    PuntoDeMenuVisible Me.mnFacPedidos(13), b
    'QUe ponga el separador
    If b Then
        If Me.mnFacPedidos(13).visible Then Me.mnFacPedidos(11).visible = True
    End If
    'Factura de mostrador se ve si el usuario lo tiene seleccionado
    'PuntoDeMenuVisible mnFacAlbMostrador, vParamAplic.FrasMostradorSerieDistinta
    
    
    
        'Si es empresa de b
    'Utilidades de traspaso presu a factura
    'y eliminar presu
    b = False
    'If vParamAplic.AlmacenB > 90 Then
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.codigo Mod 1000 = 0 Then
            b = True
        Else
            b = Val(vUsu.AlmacenPorDefecto2) > 90
        End If
    End If
    PuntoDeMenuVisible mnUtilidadesVarias(2), b
    PuntoDeMenuVisible mnUtilidadesVarias(3), b
    
    
    
    
    
    
    'De momento no hay
    mnComCtrlAlb(2).visible = False
    
    'Facturacion electronica
    PuntoDeMenuVisible mnEnvioFactuasMail(1), vParamAplic.PathFacturaE <> ""
    
    'Tratamientos
    'PuntoDeMenuVisible mnTratamientosRaiz, vParamAplic.Ariagro <> "" 'si tiene conexion ariagro ALZIRA
    PuntoDeMenuVisible mnTratamientosRaiz, vParamAplic.LlevaADV  'si tiene conexion ariagro ALZIRA
    
    
    PuntoDeMenuVisible Me.mnTratamientos(7), False  'Oculto para futuros usos
    
    'Renting y servicios
    PuntoDeMenuVisible Me.mnManPrevFac2(2), vParamAplic.Renting
    PuntoDeMenuVisible Me.mnManPrevFac2(3), vParamAplic.Renting
    PuntoDeMenuVisible Me.mnManPrevFac2(4), vParamAplic.Renting
    If vParamAplic.Renting Then
        mnManPrevFac2(4).Caption = "Facturación " & RentingLB & " y servicios"
        mnManPrevFac2(3).Caption = "Previsión " & LCase(mnManPrevFac2(4).Caption)
        
    End If
    
    PuntoDeMenuVisible mnProcesoLiquidacionProveedores, False
    
    'Tfnos x cliente
    PuntoDeMenuVisible Me.mnFacInformesVarios(6), vParamAplic.TieneTelefonia2 > 0
    PuntoDeMenuVisible Me.mnFacInformesVarios(7), vParamAplic.TieneTelefonia2 > 0
    
    PuntoDeMenuVisible Me.mnFacClientesPot, vParamAplic.ClientesPotenciales
    
    
    'Comunicacion datos ALMAGRUPO
    PuntoDeMenuVisible mnTelematel(1), vParamAplic.ComunicaAlmagrupo
    
    PuntoDeMenuVisible mnAgua, vParamAplic.AguasPotables
    
    PuntoDeMenuVisible Me.mnHuertos, vParamAplic.Huertos
    
      
    'ELUER
    b = False
    If vParamAplic.NumeroInstalacion = 4 Then b = True
    
    PuntoDeMenuVisible mnUtilidadesVarias(4), b 'vParamAplic.NumeroInstalacion = 4
    PuntoDeMenuVisible mnMtoEuler(0), b
    PuntoDeMenuVisible mnMtoEuler(1), b
    PuntoDeMenuVisible mnMtoEuler(2), b
    PuntoDeMenuVisible mnMtoEuler(3), b
    If vParamAplic.NumeroInstalacion = 4 Then
        PuntoDeMenuVisible mnRepAlbaranes, False
        PuntoDeMenuVisible mnRepEntReparacion, False
        
        
        
        'Quito los de mantenimientos
        '
        PuntoDeMenuVisible mnRepEntReparacion, False
        PuntoDeMenuVisible mnRepControlRep, False
        PuntoDeMenuVisible mnRepNumSerie, False
        PuntoDeMenuVisible mnRepMotivosBaja, False
        PuntoDeMenuVisible mnRepMotivosPend, False
        PuntoDeMenuVisible mnRepHistorico, False
        PuntoDeMenuVisible mnManServicioAsisTecn, False
        PuntoDeMenuVisible mnTiposAveria, False
        PuntoDeMenuVisible mnTrabaRealiz, False
        PuntoDeMenuVisible mnMtoEuler(0), False
    End If
    
    
    
    
    'Declaracion  fitosnaitarios
    mnUtiDeclaraLOM(0).visible = vParamAplic.LotesGeneralitat 'SUBVENCIONADOS
    mnUtiDeclaraLOM(1).visible = vParamAplic.ManipuladorFitosanitarios2
    
    
    
    'Lo pong aqui el 11 de Enero de 2011
    'Comprobar que los iconos de la barra su correspondiente
    'entrada de menu esta habilitada sino desabilitar
    PoneBarraMenus2
    
    'NUEVO 2017
    'Contabilizacion
    ComprobarFechaContabilizadas
    
    '--
    Screen.MousePointer = vbDefault
End Sub




Private Sub PuntoDeMenuVisible(ByRef MnPuntoDMenu As Menu, b As Boolean)
    If MnPuntoDMenu.visible Then MnPuntoDMenu.visible = b
    
End Sub




Private Sub MDIForm_Load()
'Formulario Principal

    CargaImagen

    PrimeraVez = True
    'Botones
    With Me.Toolbar1
        .ImageList = Me.ImgListPpal
        .Buttons(1).Image = 1   'Articulos
        .Buttons(2).Image = 2   'Movimientos Articulos
        
        .Buttons(5).Image = 3   'Clientes
        .Buttons(6).Image = 4   'Proveedores

        .Buttons(9).Image = 5   'Ofertas Clientes
        .Buttons(10).Image = 6   'Pedidos Clientes
        .Buttons(11).Image = 7   'Albaranes Clientes
        .Buttons(12).Image = 8   'Hist. Albaranes Clientes (Facturas)

        .Buttons(13).Image = 34   'FACTURAS MOSTRADOR

        .Buttons(15).Image = 9   'Pedidos Proveedor
        .Buttons(16).Image = 10   'Albaranes Proveedor
        .Buttons(17).Image = 11   'Facturas Proveedor
        .Buttons(18).Image = 12   'Recepcion Facturas Proveedor
        
        .Buttons(21).Image = 15   'Mantenimientos
        
        .Buttons(22).Image = 16   'Nº Serie
        'Si tiene PARTES seran partes
        If vParamAplic.Ariagro = "" Then
            .Buttons(23).Image = 23   'Avisos
            .Buttons(23).ToolTipText = "Aviso reparacion"
        Else
            .Buttons(23).Image = 35   'Partes trabajo
            .Buttons(23).ToolTipText = "Partes de trabajo"
        End If
        
        
        If vParamAplic.NumeroInstalacion <> 4 Then
            .Buttons(24).Image = 13 'Gastos tecnicos
            .Buttons(24).ToolTipText = "Gastos ténicos"
        Else
            .Buttons(24).Image = 37 'Gastos tecnicos
            .Buttons(24).ToolTipText = "Reloj"
        End If
        
        
        
        .Buttons(25).Image = 22 'Consulta precio articulo
        .Buttons(26).Image = 19 'Pantalla venta del TPV
        .Buttons(27).Image = 21 'Agenda
        .Buttons(28).Image = 20 'Agenda
        
        .Buttons(30).Image = 14 'Salir
    End With
    LeerEditorMenus
    PonerDatosFormulario False
    
       
    'Fijar primer dia la semana en vbMyMonday
    'Para el calendario.
    FijarPrimerDiaSemana
    
    
    If vParamAplic.NumeroInstalacion = 4 Then
        Me.WindowState = 0
        Me.Width = Screen.Width - 1200
        Me.Height = Screen.Height - 3000

    Else
        Me.WindowState = 2
    End If
    
       
    
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\arifon2.dat")
    If Err.Number <> 0 Then
        Me.Picture = LoadPicture()
        Err.Clear
    End If
End Sub




Private Sub PonerDatosFormulario(DesdeCambiarEmpresa As Boolean)
Dim Config As Boolean


    If Not DesdeCambiarEmpresa Then
        Config = (vEmpresa Is Nothing) Or (vParam Is Nothing) Or (vParamAplic Is Nothing)
    
        If Config Then HabilitarSoloPrametros_o_Empresas False
    End If
    
    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If DesdeCambiarEmpresa Then
        ReestablecerMenus
        HabilitarSoloPrametros_o_Empresas True
    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    

    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'Formulario Principal
Dim cad As String


    'Elimnar bloquo BD
    Set vUsu = Nothing
    Set vConfig = Nothing
    Set vEmpresa = Nothing
    
    Set vParam = Nothing
    Set vParamAplic = Nothing
    
    
    TerminaBloquear
    
    'cerrar las conexiones
    conn.Close
    CerrarConexionConta

End Sub




Private Sub mnAdmAgen2_Click(Index As Integer)
    Select Case Index
    Case 0
        'Ventas x agente
        AbrirListado2 36
        
    
    Case 1
        
        'beneficio por agente
        AbrirListado2 37
    
    Case 2
'        frmListado3.Opcion = 37    'LISTADO 3
'        frmListado3.Show vbModal
        AbrirListado3 37
    Case 4
        frmFacComisionAgen.Show vbModal
    
    Case 5
'        frmListado3.Opcion = 31
'        frmListado3.Show vbModal
        AbrirListado3 31
    Case 7
        'Comisiones ECO
        AbrirListado3 43
    Case 8
        AbrirListado3 46
    Case 9
        AbrirListado2 49
    End Select
End Sub



Private Sub mnAdmGastosTec_Click()
'Gastos Técnicos
    frmAdmGasTec.Show vbModal
End Sub

Private Sub mnAdministra_Click(Index As Integer)
    Select Case Index
    Case 1
        'Caluclo de riesgo
        If vUsu.Nivel > 0 Then
            MsgBox "No tiene permiso", vbExclamation
        Else
            AbrirListado2 31
        End If
        
    Case 2
'        frmListado3.Opcion = 25
'        frmListado3.Show vbModal
        AbrirListado3 25
    Case 3
        'Informe prevision de tesoreria
'        frmListado3.Opcion = 1
'        frmListado3.Show vbModal
        AbrirListado3 1
    Case 4
        'MsgBox "Comming soon men!!!!", vbExclamation
        'Correcion de costes de articulos varios en facturas
        frmFacCosteLin.Show vbModal
        
    Case 5
        'Modificar coste estadistica ventas
'        frmListado3.Opcion = 11
'        frmListado3.Show vbModal
        AbrirListado3 11
    Case 6
        'FLOTAS. Despliega submenu
    Case 7
        AbrirListado3 40
        
    End Select
        
End Sub

Private Sub mnAdmNominas_Click()
'Nominas y Gastos
    frmAdmNominas.Show vbModal
End Sub

Private Sub mnAdmTrabajadores_Click()
    frmAdmTrabajadores.Show vbModal
End Sub

Private Sub mnAgenda_Click()
    'MsgBox "Se ha producido un error abriendo la agenda", vbExclamation
    'FALTA###
    'frmMainCalendar.Show
    
    MsgBox "Avise soporte técnico. Falta OCX Codejock", vbCritical
    
    
End Sub

Private Sub mnAguaLin_Click(Index As Integer)

    Select Case Index
    Case 0
        frmAguaContadores.Show vbModal
    
    Case 1
        AbrirListado3 52
        
    Case 2
        AbrirListado3 51
    
    
    Case 3
        'Resumen facturacion 53
        AbrirListado3 53
        
        
    Case 4
        'Listado para rellenar modelos 100,101,102 EPSAR
        'de facturaciones canon generalitat
        AbrirListado3 55
    Case 5
        'Listado exportacion contadores
        AbrirListado3 60
        
    Case 6
        frmListado4.Opcion = 13
        frmListado4.Show vbModal
        
    Case 7
        'Declaracion detallada ejereccio
        AbrirListado3 58
    
    Case 8
        frmAguaCalibres.Show vbModal
    Case 10
        frmAguaParam.Show vbModal
    End Select
End Sub

Private Sub mnAlbaranesB_Click()
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALZ"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = "ALZ"
        frmFacEntAlbSAIL.EsHistorico = False
        frmFacEntAlbSAIL.Show vbModal
    End If
End Sub



Private Sub mnAlmActualizarInve_Click()
    AbrirListado (14)
End Sub

Private Sub mnAlmAlPropios_Click()
    frmAlmAlPropios.Show vbModal
End Sub

Private Sub mnAlmArticulos_Click()
    frmAlmArticulos.DatosADevolverBusqueda = ""
    frmAlmArticulos.Show vbModal
End Sub


Private Sub mnAlmCategoria_Click()
    'categorias de articulos
    frmAlmCategorias.Show vbModal
End Sub

Private Sub mnAlmControlStockDesdeInv_Click()
    'frmListado3.Opcion = 27
    'frmListado3.Show vbModal
    AbrirListado3 27
End Sub

Private Sub mnAlmEntradaInve_Click()
    frmAlmInventario.Show vbModal
End Sub

Private Sub mnAlmFamiliaArticulo_Click()
    frmAlmFamiliaArticulo.Show vbModal
End Sub


Private Sub mnAlmHcoInven_Click()
    frmAlmHcoInven.Show vbModal
End Sub

Private Sub mnAlmListadoInve_Click()
    AbrirListado (13)
End Sub

Private Sub mnAlmListadosVarios_Click(Index As Integer)
    
    

    Select Case Index
    Case 0
            'Informe de Stocks a una Fecha
            AbrirListado (19)
    Case 1
            'Stocks por meses
            'frmListado3.Opcion = 4
            'frmListado3.Show vbModal
            AbrirListado3 4
    Case 2
            'frmListado3.Opcion = 26
            'frmListado3.Show vbModal
            AbrirListado3 26
    Case 3
            'frmListado3.Opcion = 35
            'frmListado3.Show vbModal
            AbrirListado3 35
    Case 4
            'Informe stock minimo
            AbrirListado (100)
    End Select
End Sub

Private Sub mnAlmListComponentes_Click()
'Informe de articulos q estan compuestos de otros articulos
    AbrirListado (11)
End Sub

Private Sub mnAlmListInactivos_Click()
    AbrirListado (15)
End Sub

Private Sub mnAlmListMaxMin_Click()
'Informe de Stocks Maximos y Minimos
    AbrirListado (18)
End Sub

Private Sub mnAlmListMovim_Click()
    AbrirListado (9)
End Sub

Private Sub mnAlmListValoracion_Click()
    AbrirListado (17)
End Sub

Private Sub mnAlmMarcas_Click()
    frmAlmMarcas.Show vbModal
End Sub

Private Sub mnAlmMovimArticulos_Click()
    frmAlmMovimArticulos.Show vbModal
End Sub

Private Sub mnAlmMovimArticulosSt_Click()
    frmAlmMovArtSaldo.Show vbModal
End Sub

Private Sub mnAlmMovimientos_Click()
    frmAlmMovimientos.EsHistorico = False
    frmAlmMovimientos.hcoCodMovim = -1 'No carga el form al abrir
    frmAlmMovimientos.Show vbModal
End Sub

Private Sub mnAlmMovimientosHco_Click()
    frmAlmMovimientos.EsHistorico = True
    frmAlmMovimientos.hcoCodMovim = -1
    frmAlmMovimientos.Show vbModal
End Sub

Private Sub mnAlmNumLotes_Click()
'numero de lote de los artículos
    frmAlmNumLote.Show vbModal
End Sub





Private Sub mnAlmTipoUnidad_Click()
    frmAlmTipoUnidad.Show vbModal
End Sub

Private Sub mnAlmTomaInven_Click()
    AbrirListado (12)
End Sub

Private Sub mnAlmTraspaso_Click()
    frmAlmTraspaso.EsHistorico = False
    frmAlmTraspaso.hcoCodMovim = -1
    frmAlmTraspaso.Show vbModal
End Sub

Private Sub mnAlmTraspasoHco_Click()
    frmAlmTraspaso.EsHistorico = True
    frmAlmTraspaso.hcoCodMovim = -1
    frmAlmTraspaso.Show vbModal
End Sub

Private Sub mnAlmUbicacion_Click()
    frmAlmUbicaciones.Show vbModal
End Sub

Private Sub mnAlmValoracionInve_Click()
    AbrirListado (16)
End Sub



Private Sub mnAridoc1_Click(Index As Integer)


    'Configuracion aridoc
    If Index = 1 Then HacerMenuARidoc 0
    
End Sub

Private Sub mnAridocFacturas_Click()
    frmAridocSeleccion.vOpcion = 1
    frmAridocSeleccion.Show vbModal
End Sub

Private Sub mnArticulos2_Click(Index As Integer)
    Select Case Index
    Case 0
        frmVarios.Opcion = 1
        frmVarios.Show vbModal
    Case 1
        AbrirListado3 49
    Case 2
        'If vUsu.Nivel > 0 Then Exit Sub
        
        'Bloquear proceso
        If BloqueoManual("CambioArt", "1") Then
            frmAlmCambRef.Show vbModal
            DesBloqueoManual "CambioArt"
        End If
        
        
        
        
    End Select
End Sub

Private Sub mnBackUp_Click()
'Copia de seguridad de toda la base de datos
    frmBackUP.Show vbModal
End Sub

Private Sub mnBorrarAvisosCerrados_Click()
    AbrirListado 83
End Sub




Private Sub mnCambioEmpresa_Click()
    
    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If

    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.codigo


    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"
    
    frmLogin.Show vbModal

    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    conn.Close
    ConnConta.Close


    'Abre la conexión a BDatos:Ariges
    If AbrirConexion() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    Else
        Set vParam = Nothing
        Set vParamAplic = Nothing
        'Carga Parametros Generales y Contables de la empresa
        LeerParametros
    End If


    'Abrir conexión a la BDatos de Contabilidad para acceder a
    'Tablas: Cuentas, Tipos IVA
    If AbrirConexionConta(False) = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
        End
    End If

    
    
    Set vEmpresa = Nothing
    'LeerEmpresaParametros
    
     'Carga los Datos Básicos de la empresa
    LeerDatosEmpresa
    
    
    'Carga los Niveles de cuentas de Contabilidad de la empresa
    LeerNivelesEmpresa
    
    
    
    
    If vParamAplic.QueEmpresaEs = 2 Then
        Me.Hide
        frmPpalGessocial.Show vbModal
        Me.Show
        mnCambioEmpresa_Click
        Exit Sub
    End If
    
    
    
    
    
'    PonerDatosFormulario
    PonerDatosVisiblesForm

    'Ponemos primera vez a false
    PonerDatosFormulario True
    PrimeraVez = True
    MDIForm_Activate

    

    Screen.MousePointer = vbDefault
End Sub


Private Sub mnCambioPwd_Click()
    frmListado4.Opcion = 4
    frmListado4.Show vbModal
End Sub

Private Sub mnCartaRenovaMante_Click()
    AbrirListado 78
End Sub

'Private Sub mnCheckVersion_Click()
''    Screen.MousePointer = vbHourglass
''    LanzaHome "webversion"
''    espera 2
''    Screen.MousePointer = vbDefault
'End Sub


Private Sub mnComAlbMan_Click()
    'Mantenimiento de Albaranes a Proveedor
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmComEntAlbaranes.EsHistorico = False
        frmComEntAlbaranes.Show vbModal
    Else
        frmComEntAlbaranSA.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmComEntAlbaranSA.EsHistorico = False
        frmComEntAlbaranSA.Show vbModal
    
    End If
End Sub



Private Sub mnComContFactu_Click()
'Contabilizar Facturas
    AbrirListado (224) 'Para pedir datos
End Sub

Private Sub mnComCtrlAlb_Click(Index As Integer)
    If Index = 1 Then
        'frmComAlbAsignar.Show vbModal
        frmComCtrDoc.Show vbModal
    End If
End Sub

Private Sub mnComDirecciones_Click()
    frmComDirecciones.Show vbModal
End Sub










Private Sub mnComEstadisticaLin_Click(Index As Integer)
    Select Case Index
    Case 0
        'Listado de compras por proveedor
        AbrirListadoOfer (310)
    Case 1
        'Listado de compras por Familia
        AbrirListadoOfer (311)
    Case 2
        frmVarios.Opcion = 11
        frmVarios.Show vbModal
    Case 3
        'Listado de alb compras por proveedor
        AbrirListadoOfer (312)
    Case 4
         'frmListado3.Opcion = 7
         'frmListado3.Show vbModal
         AbrirListado3 7
    Case 5
        AbrirListado2 50
    End Select
End Sub



Private Sub mnComFacturar_Click()
    frmComFacturar.Show vbModal
End Sub

Private Sub mnComHcoAlbaranes_Click()
'Historico albaranes de compras a proveedores
     If vParamAplic.TipoFormularioClientes = 0 Then
        frmComEntAlbaranes.EsHistorico = True
        frmComEntAlbaranes.Show vbModal
    Else
        frmComEntAlbaranSA.EsHistorico = True
        frmComEntAlbaranSA.Show vbModal
    End If
End Sub

Private Sub mnComHcoFacturas_Click()
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmComHcoFacturas2.hcoCodMovim = ""
        frmComHcoFacturas2.Show vbModal
    Else
        'SAIL
        frmComHcoFacturSA.hcoCodMovim = ""
        frmComHcoFacturSA.Show vbModal
    End If
End Sub







Private Sub mnComInfVarios1_Click(Index As Integer)
    Select Case Index
    Case 0
        'Informe de Proveedores
        AbrirListado (58)   ': Informe Proveedores
    Case 1
        'Etiquetas de proveedores
        AbrirListadoOfer (305) '305: Informe Etiquetas de Proveedores
    
    Case 2
        'Cartas a proveedores
        AbrirListadoOfer (306) '306: Informe Cartas a Proveedores
        
    Case 3
         AbrirListado 101
    
    End Select
    
End Sub

Private Sub mnComPedidosLin_Click(Index As Integer)
    'Cuelga de COMPRAS --- PEDIDOS---
    Select Case Index
    Case 0, 1
        'Mnatenimiento de Pedidos de compras
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmComEntPedidos2.MostrarDatos = ""
            frmComEntPedidos2.EsHistorico = Index = 1
            frmComEntPedidos2.Show vbModal
        Else
            'SAIL
            frmComEntPedidosSa.MostrarDatos = ""
            frmComEntPedidosSa.EsHistorico = Index = 1
            frmComEntPedidosSa.Show vbModal
        End If
    'Case 1
    '    frmComEntPedidos.MostrarDatos = ""
    '    frmComEntPedidos.EsHistorico = True
    '    frmComEntPedidos.Show vbModal
    Case 2
        'Listado de material pendiente de recibir
        AbrirListadoOfer (307) '307: List. Materia pte recibir
    Case 4
        'Propuesta de pedido
        AbrirListado2 32
    End Select
End Sub

Private Sub mnComPreProv_Click(Index As Integer)
    Select Case Index
    Case 0
        'precios proveedor
        frmComPreciosProv2.NuevoDato = "" 'Para que no se poing en modo insercion
        frmComPreciosProv2.Show vbModal
    Case 1
        'Dto proveedor
        frmComDtosFamMarca.Show vbModal
    
    Case 2
        'Copiar desde venta
        CadenaDesdeOtroForm = "V"
        AbrirListado2 28
    Case 3
            frmFacActPrecios2.Proveedor = True
        frmFacActPrecios2.Show vbModal
    End Select
End Sub

Private Sub mnComProveedores_Click()
'Compras. Mantenimiento de Proveedores
    frmComProveedores.Show vbModal
End Sub


Private Sub mnComProveVarios_Click()
'Proveedores varios
    frmComProveV.Show vbModal
End Sub

Private Sub mnComPteFacturar_Click()
'Listado de Albaranes pendientes de Factura
    AbrirListadoOfer (308) '308: List. Albaranes pte facturar
End Sub



Private Sub mnConfManteUsuarios_Click()
'Mantenimiento de Usuarios

      frmMantenusu.Show vbModal
      
End Sub

Private Sub mnConfParamAplic_Click()
'Parametros de la Aplicación
    Screen.MousePointer = vbHourglass
    Load frmConfParamAplic
    frmConfParamAplic.Show vbModal
    
End Sub



Private Sub mnConfParamGenerales_Click()
'Parametros generales de la Empresa

    frmConfParamGral.Show vbModal
End Sub



Private Sub mnConfParamRpt_Click()
'Parametros para informes de Crystal Report
    frmConfParamRpt.Show vbModal
End Sub

Private Sub mnConTMovimiento_Click()
'Mantenimientos de los tipos de movimientos
    frmConfTipoMov.Show vbModal
End Sub


Private Sub mnCRM_Click(Index As Integer)
    
        Select Case Index
        Case 0
            frmCRMMto.DesdeElCliente = 0 'No clien
            frmCRMMto.TipoPredefinido = 0   'Ninguno
            frmCRMMto.Show vbModal
            
        Case 1
            frmCRMtipos.Show vbModal
        
        Case 2
            frmCRMVarios.Opcion = 0
            frmCRMVarios.Show vbModal
            
        Case 3
            frmListadoOfer.OpcionListado = 406
            frmListadoOfer.Show vbModal
        Case 4
            frmCRMVarios.Opcion = 1
            frmCRMVarios.Show vbModal
            
        End Select
        
End Sub

Private Sub mnDtoCantidad_Click()
    frmFacDtoUd.Show vbModal
End Sub



Private Sub mnEliminarFacturas_Click()
    AbrirListado 97
End Sub

Private Sub mnEnvioFactuasMail_Click(Index As Integer)
    AbrirListadoOfer 315 + Index
End Sub

Private Sub mnEstadisticaReparacionTecnico_Click()
    AbrirListado2 2
End Sub


Private Sub mnEstVtasArituclo_Click()
    'frmListado3.Opcion = 18
    'frmListado3.Show vbModal
    AbrirListado3 18
End Sub

Private Sub mnEtiqMante_Click()
    AbrirListado 79
End Sub



Private Sub mnFacActividades_Click()
    frmFacActividades.Show vbModal
End Sub

Private Sub mnFacAgentesCom_Click()
    frmFacAgentesCom.Show vbModal
End Sub

Private Sub mnFacAlbDevolucion_Click()
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "DEV"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
    End If
End Sub

Private Sub mnFacAlbMostrador_Click()
'Abre el formulario de Albaranes para introducir el Albaran de Mostrador
'y desde este generar la Factura de mostrador
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALM"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
    End If
End Sub


Private Sub mnFacAlbRectifica_Click()
'Facturas Rectificativas
    'Abre el formulario de Albaranes para introducir el Albaran Rectificativo
    'y desde este generar la Factura Rectificativa
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ART"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = "ART"
        frmFacEntAlbSAIL.EsHistorico = False
        frmFacEntAlbSAIL.Show vbModal
    End If
End Sub

Private Sub mnFacAlbxArtic_Click()
'Informe de Albaranes por Articulo
    AbrirListadoPed (49)
End Sub


Private Sub mnFacBancosPropios_Click()
    frmFacBancosPropios.Show vbModal
End Sub





Private Sub mnFacCartas_Click()
'Mantenimiento de Cartas
    frmFacCartasOferta.Show vbModal
End Sub


Private Sub mnFacClientes_Click()
'Mantenimiento de Clientes
    frmFacClientes.Show vbModal
End Sub

Private Sub mnFacClientesPot_Click()
    frmFacClienPot.Show vbModal
End Sub

Private Sub mnFacClientesV1_Click()
'Mantenimiento de Clientes Varios
    frmFacClientesV.Show vbModal
End Sub



Private Sub mnFacContFactu_Click()
'Contabilizar Facturas
    AbrirListado (223) 'Para pedir datos
End Sub






Private Sub mnFacEntAlbaran_Click()
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALV"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
        
    'ElseIf T Then
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = "ALV"
        frmFacEntAlbSAIL.EsHistorico = False
        frmFacEntAlbSAIL.Show vbModal
    End If
End Sub











Private Sub mnFacEstadistica2_Click(Index As Integer)
    Select Case Index
    Case 0
        'Por proveedor
        AbrirListado2 6
    Case 1
    
        'Ventas por agente
        AbrirListado2 16
    Case 2
        'Detalle facturacion clientes
        AbrirListadoOfer (231)
    Case 3
        'Estadistica margen ventas por artículo
        AbrirListado (246)
    Case 4
        'Vtas x tipo d precio
        AbrirListado3 38
    Case 5
        'Articulos ams vendidos
        AbrirListado3 39
    Case 6
        'ALZIRA
        frmListado4.Opcion = 7
        frmListado4.Show vbModal
    
    Case 7
        'Listado pedidos por "peticion " cliente (si-no)
        AbrirListado3 63
    End Select
End Sub

Private Sub mnFacEstVentaCliente_Click()
'Estadistica Ventas por cliente
    AbrirListadoPed (227)
    BorrarTempInformes
End Sub

Private Sub mnFacEstVentaFam_Click()
'Listado de estadistica ventas por familia de articulo
    AbrirListadoOfer (230)
End Sub

Private Sub mnFacEstVentaMes_Click()
'Estadistica Ventas por Meses
    AbrirListadoPed (229)
    
End Sub

Private Sub mnFacEstVentaTraba_Click()
'Estadistica Ventas por Trabajador
    AbrirListadoPed (228)
End Sub



Private Sub mnFacFacturarAlb_Click()
'Facturacion de Albaranes de Ventas

    If vParamAplic.TipoFormularioClientes = 0 Then

        frmListadoPed.codClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
        
    Else
        'PARA sail
        frmFacturaCliSail.Show vbModal
    End If
End Sub

Private Sub mnFacFormasPago_Click()
    frmFacFormasPago.Show vbModal
End Sub



Private Sub mnFacHcoAlbaranes_Click()
'Histórico de Albaranes eliminados
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALV"
        frmFacEntAlbaranes2.EsHistorico = True
        frmFacEntAlbaranes2.Show vbModal
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = "ALV"
        frmFacEntAlbSAIL.EsHistorico = True
        frmFacEntAlbSAIL.Show vbModal
    
    End If
End Sub

Private Sub mnFacHcoFacturas_Click()
'Histórico de Facturas
    frmFacHcoFacturas2.hcoCodMovim = ""
    frmFacHcoFacturas2.Show vbModal
End Sub


Private Sub mnFacIncidencias_Click()
    frmIncidencias.Show vbModal
End Sub

Private Sub mnFacIncumPlazos_Click()
'Incumplimiento de los Plazos de Entrega
    
    AbrirListadoPed (51)
End Sub










Private Sub mnFacInformesVarios_Click(Index As Integer)
    Select Case Index
    Case 0
        'Informe de Clientes Inactivos
        AbrirListadoOfer (46) '46: Informes Clientes Inactivos
    Case 1
        'Informe de Clientes
        AbrirListadoOfer (47) '47: Informes Clientes
    Case 2
        'Informe de Altas de Nuevos Clientes
        AbrirListadoOfer (48) '48: Informes Altas Clientes

    Case 3
        'Etiquetas de clientes
        AbrirListadoOfer (90) '90: Informe Etiquetas de Clientes
    Case 4
        'Cartas a clientes
         AbrirListadoOfer (91) '91: Informe Cartas a Clientes
    Case 5
        'Listado de etiquetas de los bultos
        AbrirListado 95
    Case 6
        'Listado de telefonos por clientes
        AbrirListado3 41
    Case 7
        'Listado de telefonos por clientes
        AbrirListado3 48
    
    End Select
    
End Sub

Private Sub mnFacOfertas_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn
    
    Select Case Index
    Case 0, 5
            'Private Sub mnFacEntOfertas_Click()
         If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntOfertas2.DatosOferta = ""
            frmFacEntOfertas2.EsHistorico = Index = 5
            frmFacEntOfertas2.Show vbModal
        Else
            frmFacEntOferSAIL.DatosOferta = ""
            frmFacEntOferSAIL.EsHistorico = Index = 5
            frmFacEntOferSAIL.Show vbModal
        End If

    Case 1
            'Private Sub mnFacGrupoPlant_Click()
            'Mantenimiento de Grupos de Plantillas
        frmFacGrupoPlantilla.Show vbModal
    
    Case 2
            'Private Sub mnFacPlantillas_Click()
            'Mantenimiento de Plantillas
        frmFacPlantilla.Show vbModal
        
    Case 3
            ' Private Sub mnFacOfeEfectuadas_Click()
            'Listado de Ofertas Efectuadas
        AbrirListadoOfer (34) '34: Informe Ofertas Efectuadas
    
        
        
    'case 4  'Es la barra separadora
    
    Case 6
        
            'Private Sub mnFacTrasHist_Click()
            'Traspaso de Ofertas a las tablas de Historico
        frmListadoOfer.OpcionListado = 36
        AbrirListadoOfer (36) 'NO IMPRIME LISTADO, hace traspaso de Ofertas de la tabla (scapre) a (schpre)

    
    End Select
End Sub

Private Sub mnFacPedidos_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn
  
    Select Case Index
    Case 0, 1
        'Mantenimiento de Pedidos  Y Histórico de Pedidos
        
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntPedidos.EsHistorico = Index = 1
            frmFacEntPedidos.Show vbModal
        Else
            frmFacEntPedSail.EsHistorico = Index = 1
            frmFacEntPedSail.Show vbModal
        End If
    'Case 2  es la barra de separacion
    
    Case 3
        'Confirmar pedido   mnFacConfirmPed_Click
        AbrirListadoOfer (40)
        
    Case 4
        'Pedido por articulo
        'Private Sub mnFacPedidoxArtic_Click()
        'Informe de Pedidos por Articulo
        AbrirListadoPed (41)
        
    Case 5
        'Private Sub mnFacPedidoxClien_Click()
        'Informe de Pedidos por Cliente
        AbrirListadoPed (44)
        
        
    Case 6
        'Private Sub mnFacDispStock_Click()
        'Resumen de Disponibilidad de Stocks
        AbrirListadoPed (42)
    
    Case 7
        'Pedido por zona
        frmListado2.Opcion = 26
        frmListado2.Show vbModal
        
    Case 9
        'Precio cliente
        frmFacConsultaPrecios2.Fecha = Now
        frmFacConsultaPrecios2.Show vbModal
    Case 10
        'Estadisitcas de veces consultado  precio/cliente
        frmVarios.Opcion = 2
        frmVarios.Show vbModal
    Case 12
        frmFacEntPedidRMA.Show vbModal
    Case 13
        frmFacEntPedPresu.Show vbModal
    End Select
End Sub





Private Sub mnFacPreFacturar_Click()
' Previsión Facturacion de Albaranes
    frmListadoPed.codClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO
End Sub



Private Sub mnFacReImpFactu_Click()
'Reimprimir Factuas ya contabilizadas
    AbrirListadoOfer 226
End Sub

Private Sub mnFacRutas_Click()
    frmFacRutas.Show vbModal
End Sub

Private Sub mnFacSituaciones_Click()
    frmFacSituaciones.Show vbModal
End Sub












'---------------------------------------------
'
'  Unico punto de menu para las tarifas venta
Private Sub mnFacTarVen_Click(Index As Integer)
        Select Case Index
    Case 0
        'Tarifas Venta
        frmFacTarifas.Show vbModal
    Case 1
        'Listado Precios
        frmFacTarifasPrecios.Show vbModal
    Case 2
        'Precios especiales
        frmFacPreciosEspecial.CadenaSituarData = ""
        frmFacPreciosEspecial.Show vbModal
    
    Case 3
        'PROMOCIONES
        frmFacPromociones.Show vbModal
    
    Case 4
        'Dots familia marca
        frmFacDtosFamMarca.Show vbModal
    
    Case 5
        'dtos por activiad
        frmFacDtosAsignar.Show vbModal
    
    Case 6
        'Bonificacines factura
        frmFacBonificacion.Show vbModal
    
    Case 8
        'Actualizar precios actuales y especiales
        frmFacActPrecios2.Proveedor = False
        frmFacActPrecios2.Show vbModal
    
    Case 9
        'Copiar desde compra
        CadenaDesdeOtroForm = ""
        AbrirListado2 28
    Case 11
        'Informe control margenes de tarifas
        AbrirListado (245)
    
    Case 12
        'Correcion
        AbrirListado 247
    Case 13
        'frmListado3.Opcion = 13
        'frmListado3.Show vbModal
        AbrirListado3 13
    End Select
End Sub

Private Sub mnFacturarCliente_Click()
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacturacionCli.Show vbModal
    Else
        frmFacturaCliSail.ImprimirCertificacion = False
        frmFacturaCliSail.Show vbModal
    End If
End Sub

Private Sub mnFacturarPresupuestos_Click()
        frmListadoPed.codClien = "ALZ" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
End Sub

Private Sub mnFacZonas_Click()
    frmFacZonas.Show vbModal
End Sub

Private Sub mnFacFormasEnvio_Click()
    frmFacFormasEnvio.Show vbModal
End Sub

Private Sub mnFlotas_Click(Index As Integer)
    Select Case Index
    Case 0
        frmFlotaReg.DatosADevolverBusqueda = ""
        frmFlotaReg.Show vbModal
    Case 1
        frmFlotas.DatosADevolverBusqueda = ""
        frmFlotas.Show vbModal
        
    Case 2
    
        frmFlotasConceptos.DatosADevolverBusqueda = ""
        frmFlotasConceptos.Show vbModal
    End Select
End Sub

Private Sub mnFrecuencias_Click()
    frmFrecuencias.Show vbModal
End Sub

Private Sub mnHcoMaten_Click()
    frmManMantenimientosAnu.Show vbModal
End Sub


Private Sub mnHuertos1_Click(Index As Integer)
    If Index = 0 Then
        frmListado5.OpcionListado = 15
        frmListado5.Show vbModal
    Else
        AbrirListado3 61
    End If
End Sub

Private Sub mnInfManteAnulados_Click()
    AbrirListado 76
End Sub



Private Sub mnInfoAdm_Click(Index As Integer)
    Select Case Index
    Case 0
        '---------------------------
        '  CASE 0 y CASE 1 estan ahora en el submenu de Agentes dentro de administracion
        'frmListado2.Opcion = 36
        'frmListado2.Show vbModal
    Case 1
        'beneficio por agente
        'frmListado2.Opcion = 37
        'frmListado2.Show vbModal







    Case 3
        'beneficio por proveedor
        'frmListado2.Opcion = 40
        'frmListado2.Show vbModal
        AbrirListado2 40
    Case 4
        'frmListado2.Opcion = 41
        'frmListado2.Show vbModal
        AbrirListado2 41
        
    
    Case 5
        
        'Beneficio marca-agente-proveedor
        AbrirListado2 48
    Case 6
        'informe de articulos en promocion
        'frmListado3.Opcion = 5
        'frmListado3.Show vbModal
        AbrirListado3 5
    Case 7
        'informe de articulos en promocion
        'frmListado3.Opcion = 34
        'frmListado3.Show vbModal
        AbrirListado3 34
    Case 8
        'Ventas trabajaodr x dia
        'frmListado3.Opcion = 9
        'frmListado3.Show vbModal
        AbrirListado3 9
        
    Case 9
        'Listado ventas por forma de pago
        'frmListado3.Opcion = 19
        'frmListado3.Show vbModal
        AbrirListado3 19
    End Select
End Sub



Private Sub mnInfTeoMant_Click()
    AbrirListado 77
End Sub



Private Sub mnListadoReparacionesEfectuadas_Click()
    AbrirListado2 1
End Sub

Private Sub mnLlamadas_Click(Index As Integer)
    Select Case Index
    Case 0
        frmLlamadas.Show vbModal
        
    Case 1
        frmLlamadasTipo.Show vbModal
    End Select
End Sub

Private Sub mnManAltas_Click()
'Listado Altas de Mantenimientos
    AbrirListado 73
End Sub

Private Sub mnManEntrada_Click()
    frmManMantenimientos.Show vbModal
End Sub



Private Sub mnManFichas_Click()
'Listado Fichas de Mantenimientos
    AbrirListado 72
End Sub

Private Sub mnManListado_Click()
'Listados de Mantenimientos
    AbrirListado 70
End Sub



Private Sub mnManPrevFac2_Click(Index As Integer)
    Select Case Index
    Case 0
         ' Previsión Facturacion de Albaranes de Mantenimiento
         '    frmListadoPed.CodClien = "ALM" 'utilizamos esta vble para pasarle el tipo de movimiento
         AbrirListadoPed (74) 'NO IMPRIME LISTADO
    Case 1
        'Facturacion de Mantenimientos
         AbrirListadoPed (75) 'NO IMPRIME LISTADO
    
    Case 3
    
        'frmListado3.Opcion = 23
        frmListado3.OtrosDatos = ""
        'frmListado3.Show vbModal
        AbrirListado3 23
    Case 4
        'FACTURACION
        'frmListado3.Opcion = 22
        frmListado3.OtrosDatos = ""
        'frmListado3.Show vbModal
        AbrirListado3 22
    End Select
End Sub

Private Sub mnManRevisiones_Click()
'Listado Revisiones de Mantenimientos
     AbrirListado 71
End Sub

Private Sub mnManServicioAsisTecn_Click()
    frmManSat.Show vbModal
End Sub



Private Sub mnManteneLOG_Click()
    Screen.MousePointer = vbHourglass
    Load frmLog
    DoEvents
    frmLog.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnManTiposContrato_Click()
    frmManTiposContrato.Show vbModal
End Sub


Private Sub mnMtoEuler_Click(Index As Integer)
    If Index > 0 Then
        If Index = 1 Then
            frmFacEntAlbSAIL.hcoCodTipoM = "ALO"
        ElseIf Index = 2 Then
            frmFacEntAlbSAIL.hcoCodTipoM = "ALE"
        Else
            frmFacEntAlbSAIL.hcoCodTipoM = "ALR"
        End If
        frmFacEntAlbSAIL.Show vbModal
    End If
    
End Sub

Private Sub mnObra_Click(Index As Integer)
    Select Case Index
    Case 0
        frmObraCapitulo.Show vbModal
    Case 1
        frmObraActua.Show vbModal
    Case 2
        If vParamAplic.NumeroInstalacion = 4 Then
            frmEulerTrab.Show vbModal
        Else
            frmObrpartesTra.Show vbModal
        End If
    Case 3
        frmObraOT.Show vbModal
    Case 4
        frmEulerReloj.Show vbModal
    
    Case 6
        'Sept 2012
        frmObraListado.Opcion = 3
        frmObraListado.Show vbModal
    Case 8
        'Imprimir certificacion
        frmFacturaCliSail.ImprimirCertificacion = True
        frmFacturaCliSail.Show vbModal
    End Select
End Sub

Private Sub mnPortes_Click()
    frmFacPortes.Show vbModal
End Sub







Private Sub mnproduccion1_Click(Index As Integer)
    Select Case Index
    Case 0
        frmProdOrden.Show vbModal
    Case 1
        frmProdEnvas.Show vbModal
    Case 2
        frmAlmDescCostesTasas.Show vbModal
    Case 3
        frmListLotes.Show vbModal
    Case 4
        frmAlmCalidad.Show vbModal

    End Select
End Sub





Private Sub mnRecalPrSt_Click(Index As Integer)
    If vUsu.Nivel > 1 Then
        MsgBox "No tiene permiso", vbExclamation
        Exit Sub
    End If
    If Index = 0 Then
        frmListado3.Opcion = 6
    ElseIf Index = 1 Then
        frmListado3.Opcion = 20
    Else
        frmListado3.Opcion = 21
    End If
    frmListado3.Show vbModal
End Sub



Private Sub mnRectifInve_Click(Index As Integer)
    If vUsu.Nivel > 0 Then
        MsgBox "No tiene permiso", vbExclamation
    Else
        AbrirListado3 IIf(Index = 1, 61, 3)
    End If
End Sub

Private Sub mnRepAlbaranes_Click()
   ' If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALR"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
    'End If
End Sub

Private Sub mnRepAvisos_Click()
'Avisos de averias de clientes
    frmRepAvisos.Show vbModal
End Sub

Private Sub mnRepControlRep_Click()
'Control de Reparaciones (para los Tecnicos)
    frmRepEntReparaciones.EntradaEquipo = ""
    frmRepEntReparaciones.ControlRep = True
    frmRepEntReparaciones.EsHistorico = False
    frmRepEntReparaciones.Show vbModal
End Sub

Private Sub mnRepEntReparacion_Click()
'Mantenimiento de Reparaciones
    frmRepEntReparaciones.EntradaEquipo = ""
    frmRepEntReparaciones.ControlRep = False
    frmRepEntReparaciones.EsHistorico = False
    frmRepEntReparaciones.Show vbModal
End Sub

Private Sub mnRepFactAlb_Click()
'Facturacion de Albaranes de Reparacion
    frmListadoPed.codClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (52)
End Sub

Private Sub mnRepGarantprove_Click()
    frmListado2.Opcion = 30
    frmListado2.Show vbModal
End Sub

Private Sub mnRepHistorico_Click()
'Historico de las reparaciones
    frmRepEntReparaciones.EntradaEquipo = ""
    frmRepEntReparaciones.ControlRep = False
    frmRepEntReparaciones.EsHistorico = True
    frmRepEntReparaciones.Show vbModal
End Sub


Private Sub mnRepListAvisosPtes_Click()
'Listado de avisos de averias de clientes pendientes
    AbrirListado (409)
End Sub

Private Sub mnRepListFrecuen_Click()
'Listado de Frecuencia de Reparaciones
    AbrirListado (406)
End Sub

Private Sub mnRepListRepxClien_Click()
'Listado de las Reparaciones por cliente
    AbrirListado (64)
End Sub

Private Sub mnRepListRepxDia_Click()
'Listado de las Reparaciones del dia
    AbrirListado (63)
End Sub

Private Sub mnRepMotivosBaja_Click()
'Motivos baja equipos
    frmRepMotivosBaja.Show vbModal
End Sub

Private Sub mnRepMotivosPend_Click()
'Motivos Pendientes Reparar
    frmRepMotivosPend.Show vbModal
End Sub

Private Sub mnRepNumSerie_Click()
'Mantenimiento de Nºs de Serie
    frmRepNumSerie2.Show vbModal
End Sub

Private Sub mnRepPrevFact_Click()
' Previsión Facturacion de Albaranes de Reparacion
    frmListadoPed.codClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO

End Sub

Private Sub mnRevisarMultibase_Click()
    AbrirListado2 3
End Sub

'Private Sub mnPedirPwd_Click()
'Dim Anterior As Boolean
'
'    Anterior = Me.mnPedirPwd.Checked
'    vConfig.PedirPasswd = Not Anterior
'    If vConfig.Grabar = 1 Then
'        Me.mnPedirPwd.Checked = Anterior
'    Else
'        Me.mnPedirPwd.Checked = Not Anterior
'    End If
'End Sub


Private Sub mnSeleccionarImpresora_Click()
    Screen.MousePointer = vbHourglass
    Me.CommonDialog1.ShowPrinter
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnServicios_Click(Index As Integer)
    If Index = 0 Then Exit Sub  'La barra no puede
    Select Case Index
    Case 1, 3
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            If Index = 1 Then
                frmFacEntAlbaranes2.hcoCodTipoM = "ALS"
            Else
                frmFacEntAlbaranes2.hcoCodTipoM = "ALI"
            End If
            frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
            
        Else
            If vParamAplic.NumeroInstalacion = 4 Then
                
                frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
                If Index = 1 Then
                    frmFacEntAlbSAIL.hcoCodTipoM = "ALS"
                Else
                    frmFacEntAlbSAIL.hcoCodTipoM = "ALI"
                End If
                
                frmFacEntAlbSAIL.EsHistorico = False
                frmFacEntAlbSAIL.Show vbModal
            End If
        End If
    Case 2, 4
        If Index = 2 Then
            frmListadoPed.codClien = "ALS" 'utilizamos esta vble para pasarle el tipo de movimiento
        Else
            frmListadoPed.codClien = "ALI" 'utilizamos esta vble para pasarle el tipo de movimiento
        End If
        AbrirListadoPed (52)
    
    Case 5
        'Lisatado albaranes internos
        frmListado5.OpcionListado = 13
        frmListado5.Show vbModal
    End Select
    
End Sub

Private Sub mnSituaAlba_Click(Index As Integer)
    Select Case Index
    Case 0, 3
        If Index = 0 Then
            frmListado2.Opcion = 23
        Else
            frmListado2.Opcion = 27
        End If
        frmListado2.Show vbModal
    Case 1
     
        frmFacAlbAsignar.Show vbModal
    Case 2
        frmFacFacAsignar.Show vbModal
    Case 4
        
        AbrirListado3 47
        
    
    End Select
End Sub

Private Sub mnSociosProveedores_Click(Index As Integer)
    Select Case Index
    Case 0
        'Cambiar precios proveedores /socios
         AbrirListado2 7
         
    Case 1
        'Liquidacion SOCIOS
        AbrirListado2 8
        
    Case 2
        'Impresion facturas proveedores
        AbrirListado2 9
        
    Case 3
        'MsgBox "En desarrollo", vbExclamation
        'Asociar albaranes compras / vetnas
         frmComprasVentas.Show vbModal
    
    Case 4
        'listado trazabilidad
        AbrirListado2 15
    End Select
End Sub

Private Sub mnSoporte_Click(Index As Integer)
       
       

   
    Select Case Index
    Case 4
       
        Screen.MousePointer = vbHourglass
        LanzaHome ("websoporte")
        Screen.MousePointer = vbDefault

    Case 7
        'Acerca de
        Screen.MousePointer = vbDefault
        frmMensajes.OpcionMensaje = 3
        frmMensajes.Show vbModal
    End Select
    
End Sub

Private Sub mnTelefonia2_Click(Index As Integer)
    

    'ANTES 2013.
    'Ver seccion comenta de abajo
    Select Case Index
    Case 0
            frmFacEntAlbaranes2.hcoCodTipoM = "ALT"
            frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
    
    
    Case 2, 3
    
            If Index = 2 Then
                If vParamAplic.TieneTelefonia2 = 2 Then
    
'                    'importacion coarval
'                    frmTelefono1.Opcion = 5
'                    frmTelefono1.Show vbModal
'                    Exit Sub
                End If
            End If
            
            'Importar y el 4 verdatos
            If Index = 2 Then
                'Importacion
                CadenaDesdeOtroForm = ""
                frmTelefono1.Opcion = 1 'Importar
                frmTelefono1.Show vbModal
                If CadenaDesdeOtroForm = "" Then Exit Sub
                Screen.MousePointer = vbHourglass
                Espera 0.5
                DoEvents
            End If
        
            frmTelefono1.Opcion = 0
            frmTelefono1.Show vbModal



    Case 7
        frmTelBolbaite.QueOpcion = 3
        frmTelBolbaite.Show vbModal
    Case 8
        frmListado4.Opcion = 10
        frmListado4.Show vbModal
    Case 10, 11, 12, 14
    
        ' 2.- Listado descuentos comprataiivo copera
        ' 3.- Rsumen fracion
        ' 4.- Datos face
        '
        ' 6.-  Datos importados (index=!4)
        If Index = 14 Then
            frmTelefono1.Opcion = 6
        Else
            frmTelefono1.Opcion = Index - 8 '2,3,4
        End If
        frmTelefono1.Show vbModal

    
    End Select

End Sub






Private Sub mnTelefonia3_Click(Index As Integer)
    If Index = 0 Then
            '   bolbaite
            frmTelBolbaite.QueOpcion = 1
            frmTelBolbaite.Show vbModal
    Else
            frmTelDtoConsumo.Show vbModal
    End If
    
End Sub

Private Sub mnTelefonia4_Click(Index As Integer)
    If Index = 0 Then
            'bolbaite
            frmTelBolbaite.QueOpcion = 0
            frmTelBolbaite.Show vbModal
    ElseIf Index = 1 Then
         frmTelDtoCuotas.Show vbModal
    Else
         frmTelBolbaite.QueOpcion = 2
         frmTelBolbaite.Show vbModal
    End If
End Sub

Private Sub mnTelematel_Click(Index As Integer)
    Select Case Index
    Case 0
        frmTelematMto.Show vbModal
    Case 1
        frmAlmagrupo.Show vbModal
  
    End Select
End Sub

Private Sub mnTicket_Click(Index As Integer)
    
    If Index > 0 Then AbrirListado2 12 + Index

    
End Sub

Private Sub mnTiposArticulos_Click()
    frmAlmTipoArticulo.Show vbModal
End Sub

Private Sub mnSalir_Click()
    End
End Sub





Private Sub mnTiposAveria_Click()
    frmtipave.Show vbModal
End Sub


Private Sub AbirTPVpantallaVenta()
'Pantalla venta del TPV
Dim nom As String

    'Antes de abrir la pantalla de venta comprobamos que podemos leer el terminal
    'nom = ComputerNameTServer

    nom = ComputerName 'Nombre PC conectado por Terminal Server / local
    
    If Trim(nom) <> "" Then
        frmFacTPVEnt.NomrePC_conectado = nom
        frmFacTPVEnt.Show
    Else
'        'Terminal con el que trabajaremos, leemos el nombre del ordenador en local
'        'si no trabajamos en terminal server
'        nom = ComputerName
'        If Trim(nom) <> "" Then
'            frmFacTPVEnt.NomrePC_conectado = nom
'            frmFacTPVEnt.Show
'        Else
            MsgBox "No se puedo establecer un terminal.", vbExclamation
'        End If
    End If
End Sub



Private Sub mnTPVLinea_Click(Index As Integer)
            '
    Select Case Index
    Case 0
        AbirTPVpantallaVenta
    Case 1
        'Cierre caja
        'Abre el informe de cierre de caja del dia en el TPV
        AbrirListadoOfer (240)
    
    Case 2
        'Etiquetas estanteria
        AbrirListado 94
    Case 4
        'Parámetros generales del TPV
        frmFacTPVParamG.Show vbModal
    Case 5
        frmFacTPVParamT.Show vbModal
    End Select
End Sub

Private Sub mnTrabaRealiz_Click()
    frmManTraReali.Show vbModal
End Sub

Private Sub mnTraspasoMante_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.OpcionMensaje = 18
    frmMensajes.Show vbModal
End Sub





Private Sub mnTratamientos_Click(Index As Integer)
    Select Case Index
    Case 0
        'If Index = 0 Then
        frmAlmMatAct.DatosADevolverBusqueda = ""
        frmAlmMatAct.Show vbModal
    Case 1
        'If Index = 1 Then
        frmAlmADR.DatosADevolverBusqueda = ""
        frmAlmADR.Show vbModal
    Case 2
        'If Index = 2 Then
        frmAlmPlagas.DatosADevolverBusqueda = ""
        frmAlmPlagas.Show vbModal
    Case 3
        frmFlotas.Show vbModal
    Case 4
        frmADVTratamientos.DatosADevolverBusqueda = False
        frmADVTratamientos.Show vbModal
    Case 5
        frmADVTraPartes.Show vbModal
        
    Case 6
        'Fitos por campos
        frmListado5.OpcionListado = 12
        frmListado5.Show vbModal
    Case 7
        'Vacio y NO visible
        
    Case 9
        frmListado5.OpcionListado = 9
        frmListado5.Show vbModal
    End Select
End Sub

Private Sub mnUtiBuscarErrConCli_Click()
'Facturas pendientes de contabilizar (CLIENTES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 6
    frmUtilidades.Show vbModal
End Sub

Private Sub mnUtiBuscarErrConPro_Click()
'Facturas pendientes de contabilizar (PROVEEDORES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 7
    frmUtilidades.Show vbModal
End Sub


Private Sub mnUtiBuscarErrFac_Click()
'Buscar errores en nº de factura (solo en facturas de clientes)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 5
    frmUtilidades.Show vbModal
End Sub



Private Sub mnUtiConnActivas_Click()
'ver las conexiones a donde apuntan
Dim cad As String
'    cad = "Conexiones:" & vbCrLf
'    cad = cad & "------------------" & vbCrLf & vbCrLf
'    cad = cad & "Ariges: " & vbCrLf & conn.ConnectionString & vbCrLf & vbCrLf
'    cad = cad & "Conta: " & vbCrLf & ConnConta.ConnectionString & vbCrLf
'    MsgBox cad, vbInformation
    
    
    MostrarCadenasConexion
End Sub



Private Sub mnUtiDeclaraLOM_Click(Index As Integer)
    If Index = 0 Then
        frmFacLotesGeneralitat.Show vbModal
    Else
        frmUtDeclara.Show vbModal
    End If
End Sub

Private Sub mnUtilidadesVarias_Click(Index As Integer)
    Select Case Index
    Case 0
        AbrirListado3 16
    Case 1
        'Comprobar cuenta banco secciones(y contabilidades)
        
        frmListadoOfer.OpcionListado = 408
        frmListadoOfer.Show vbModal
        
    Case 2, 3
        If vUsu.Nivel > 1 Then
            MsgBox "No tienen permiso para realizar esta accion", vbExclamation
        Else
            If Index = 2 Then
                frmListado3.Opcion = 45
                frmListado3.Show vbModal
            Else
                frmListado5.OpcionListado = 2
                frmListado5.Show vbModal
            End If
        End If
        
    Case 4
        frmEulerPrecios.Show vbModal
    
    Case 5
        frmListado3.Opcion = 59
        frmListado3.Show vbModal
    Case 6
        frmListado3.Opcion = 66
        frmListado3.Show vbModal
    End Select
End Sub







Private Sub mnUtiMensLin_Click(Index As Integer)
    'Nuevo mensaje en la utilidad de mensajeria interna
    Select Case Index
    Case 0
        frmMensaje2.Show vbModal
    Case 1
    
    
    Case 3
         frmTiposMensajes.Show vbModal
    End Select
    
End Sub

Private Sub mnUtiUsuActivos_Click()
'Muestra si hay otros usuarios conectados a la Gestion
Dim SQL As String
Dim i As Integer

    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    If CadenaDesdeOtroForm <> "" Then
        i = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            SQL = RecuperaValor(CadenaDesdeOtroForm, i)
            If SQL <> "" Then Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
            i = i + 1
        Loop Until SQL = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningun usuario, además de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    CadenaDesdeOtroForm = ""
End Sub





Private Sub mnVerAvisos_Click()
    If TieneAvisosPendientes Then
        frmAlertas.Show vbModal
    Else
        MsgBox "No hay avisos para mostrar", vbInformation
    End If
End Sub







Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
            
            
     'QUitar.
     ' Pruebas para ver la semana de pedido para RAMON
'    Dim I As Integer
'    Dim F As Date
'    Dim Cad As String
'    Dim Ax As String
'
'
'    For I = 2010 To 2025
'        Ax = I & "       "
'        Ax = Ax & "  " & Format(Format("01/01/" & I, "ww", vbMonday, vbFirstJan1), "00")
'        Ax = Ax & "  " & Format(Format("01/01/" & I, "ww", vbMonday, vbFirstFourDays), "00")
'        Ax = Ax & "  " & Format(Format("01/01/" & I, "ww", vbMonday, vbFirstFullWeek), "00")
'
'
'
'        Ax = Ax & "      "
'        Ax = Ax & "  " & Format(Format("31/12/" & I, "ww", vbMonday, vbFirstJan1), "00")
'        Ax = Ax & "  " & Format(Format("31/12/" & I, "ww", vbMonday, vbFirstFourDays), "00")
'        Ax = Ax & "  " & Format(Format("31/12/" & I, "ww", vbMonday, vbFirstFullWeek), "00")
'
'        Debug.Print Ax
'    Next I
'
'
'    Stop
'
            

        
    Select Case Button.Index
    Case 1 'Mantenimiento de Artículos
        mnAlmArticulos_Click
    Case 2 'Movimientos Articulos
        mnAlmMovimArticulos_Click
        
    Case 5 'Mantenimiento Clientes
        mnFacClientes_Click
    Case 6 'Mantenimiento Proveedores
        mnComProveedores_Click
        
    Case 9 'Ofertas a Clientes
        mnFacOfertas_Click 0
    Case 10 'Pedidos a Clientes
        'mnFacEntPedidos_Click
        mnFacPedidos_Click 0
    Case 11 'Albaranes a Clientes
    
        mnFacEntAlbaran_Click
        
        
    Case 12 'Hist. Albaranes (Facturas)
        mnFacHcoFacturas_Click
        
    Case 13
        'Facturas mostrador
        mnFacAlbMostrador_Click
        
    Case 15 'Pedidos de Proveedores
        mnComPedidosLin_Click 0
    Case 16 'Albaranes de Proveedores
        mnComAlbMan_Click
    Case 17 'Facturas de Proveedores
        mnComHcoFacturas_Click
    Case 18 'Recepcion Fact. Prove
        If Me.mnComFacturar.visible And Me.mnComFacturar.Enabled Then mnComFacturar_Click
        
    Case 21 'Mantenimientos
        mnManEntrada_Click
    Case 22 'Nº Serie
        mnRepNumSerie_Click
    Case 23
        If vParamAplic.Ariagro = "" Then
            mnRepAvisos_Click
        Else
            mnTratamientos_Click 5
        End If
    Case 24 'Gastos Técnicos
    
        If vParamAplic.NumeroInstalacion = 4 Then
            mnObra_Click 4
        Else
            mnAdmGastosTec_Click
        End If
    Case 25
        'Consulta precio articulo
        mnFacPedidos_Click 9
        
    Case 26 'Entrada al TPV
        AbirTPVpantallaVenta
    Case 27
        'cambiar empresa
        mnCambioEmpresa_Click
        
    Case 28
        If vParamAplic.Frecuencias Then
            'Llamamos a frecuencias
            mnFrecuencias_Click
        Else
            mnAgenda_Click
        End If
        
    Case 30 'Salir
        mnSalir_Click
    End Select
End Sub


Private Sub PonerDatosVisiblesForm()
'Escribe texto de la barra de la aplicación
Dim cad As String
    cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    cad = cad & ", " & Format(Now, "d")
    cad = cad & " de " & Format(Now, "mmmm")
    cad = cad & " de " & Format(Now, "yyyy")
    cad = "    " & cad & "    "
    Me.StatusBar1.Panels(5).Text = cad
    If vEmpresa Is Nothing Then
        Caption = "ARIGES" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
        'Panel con el nombre de la empresa
        Me.StatusBar1.Panels(2).Text = "Falta configurar"
    Else
        Caption = "ARIGES" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vEmpresa.nomempre & "  -    Usuario: " & vUsu.Nombre
        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
    End If
End Sub


Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim cad As String

    
    For Each T In Me
        cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            If LCase(Mid(T.Caption, 1, 1)) <> "-" Then T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnConfParamAplic = True
    Me.mnConfParamGenerales = True

    Me.mnSalir.Enabled = True
    Me.mnCambioEmpresa.Enabled = True
End Sub

'-------------------------------------
'Pondremos todos los que menus a visibles. Y luego ya , en f
Private Sub ReestablecerMenus()
Dim T
      For Each T In Me
        If Mid(T.Name, 1, 2) = "mn" Then T.visible = True
    Next
End Sub

Private Sub PonerMenusNivelUsuario()
Dim b As Boolean

    b = (vUsu.Nivel = 0) Or (vUsu.Nivel = 1)  'Administradores y root

    Me.mnConfParamAplic = b
    mnConfManteUsuarios = b
    
    mnUsuarios.Enabled = b
    mnNuevaEmpresa.Enabled = b
    mnPedirPwd.Enabled = b
    Me.mnUtiConnActivas.Enabled = (vUsu.Nivel = 0) 'solo para root
    

    b = vUsu.Nivel = 3  'Es usuario de consultas
    If b Then
        'Inventario
        Me.mnAlmTomaInven.Enabled = False
        Me.mnAlmEntradaInve.Enabled = False
        Me.mnAlmActualizarInve.Enabled = False
        Me.mnAlmListadoInve.Enabled = False
        Me.mnAlmValoracionInve.Enabled = False
        'Antes
        'Me.mnFacTrasHist.Enabled = False
        mnFacOfertas(6).Enabled = False
        
        
        'Facturacion Ventas
        Me.mnFacFacturarAlb.Enabled = False
        Me.mnFacContFactu.Enabled = False
        
        'Facturacion Compras
        Me.mnComFacturar.Enabled = False
        Me.mnComContFactu.Enabled = False
        
        'Reparaciones
        Me.mnRepFactAlb.Enabled = False
        
        'Mantenimientos y renting
        'Me.mnManFactAlb.Enabled = False
        mnManPrevFac2(1).Enabled = False
        mnManPrevFac2(4).Enabled = False
    End If
End Sub



Private Sub LanzaHome(Opcion As String)
Dim i As Integer
Dim cad As String

    On Error GoTo ELanzaHome

'    LanzaHome = False
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", Opcion, "codigo", "1", "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en Parámetros de la Aplicación.", vbExclamation
        Exit Sub
    End If

    If Opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision


'    I = FreeFile
'    cad = ""
'    Open App.Path & "\lanzaexp.dat" For Input As #I
'    Line Input #I, cad
'    Close #I

    'Lanzamos
    If LanzaHomeGnral(CadenaDesdeOtroForm) Then Espera 2
    
'    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'    If vConfig.Explorador <> "" Then
'        Shell vConfig.Explorador & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'        LanzaHome = True
'    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub



Private Sub LeerEditorMenus()
Dim SQL As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from usuarios.appmenus where aplicacion='Ariges'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from usuarios.appmenususuario where aplicacion='Ariges' and codusu = " & Val(Right(CStr(vUsu.codigo), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3)
            If Right(miRsAux.Fields(3), 1) <> "|" Then SQL = SQL & "|"
            SQL = SQL & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "·" & SQL
        For Each T In Me.Controls
            If TypeOf T Is Menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                Debug.Print C
                If InStr(1, SQL, C) > 0 Then
                    
                    'Stop
                    T.visible = False
                End If
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index & "|"
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function



Private Sub PoneBarraMenus2()
'Para cada boton de la toolbar comprobar que el menu con el que se corresponde
'esta visible y activado, y ponerle los mismos valore que tenga el menu
Dim Activado As Boolean

    On Error GoTo 0
    
    '-----------------------------------------------------------
    'Articulos
    Me.Toolbar1.Buttons(1).visible = ComprobarBotonMenuVisible(Me.mnAlmArticulos, Activado)
    Me.Toolbar1.Buttons(1).Enabled = Activado

    'Movimientos de Articulos
    Me.Toolbar1.Buttons(2).visible = ComprobarBotonMenuVisible(Me.mnAlmMovimArticulos, Activado)
    Me.Toolbar1.Buttons(2).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Clientes
    Me.Toolbar1.Buttons(5).visible = ComprobarBotonMenuVisible(Me.mnFacClientes, Activado)
    Me.Toolbar1.Buttons(5).Enabled = Activado
    
    'Proveedores
    Me.Toolbar1.Buttons(6).visible = ComprobarBotonMenuVisible(Me.mnComProveedores, Activado)
    Me.Toolbar1.Buttons(6).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Ofertas Clientes
    Me.Toolbar1.Buttons(9).visible = ComprobarBotonMenuVisible(Me.mnFacOfertas(0), Activado)
    Me.Toolbar1.Buttons(9).Enabled = Activado
    
    'Pedidos Clientes
    Me.Toolbar1.Buttons(10).visible = ComprobarBotonMenuVisible(mnFacPedidos(0), Activado)
    Me.Toolbar1.Buttons(10).Enabled = Activado
    
    'Albaranes Clientes
    Me.Toolbar1.Buttons(11).visible = ComprobarBotonMenuVisible(Me.mnFacEntAlbaran, Activado)
    Me.Toolbar1.Buttons(11).Enabled = Activado
    
    'Facturas Clientes
    Me.Toolbar1.Buttons(12).visible = ComprobarBotonMenuVisible(Me.mnFacHcoFacturas, Activado)
    Me.Toolbar1.Buttons(12).Enabled = Activado
    
    
    
    'Si esta visible entonces SI lleva la misma serie no la muestro
    If vParamAplic.FrasMostradorSerieDistinta Then
        Me.Toolbar1.Buttons(13).visible = ComprobarBotonMenuVisible(mnFacAlbMostrador, Activado)
        Me.Toolbar1.Buttons(13).Enabled = Activado
    Else
        Me.Toolbar1.Buttons(13).visible = False
    End If
    
    '-----------------------------------------------------------
    'Pedidos Proveedor
    'Comprobar que los menus del que cuelga no este bloqueado o invisible
    Me.Toolbar1.Buttons(15).visible = ComprobarBotonMenuVisible(mnComPedidosLin(0), Activado)
    Me.Toolbar1.Buttons(15).Enabled = Activado
    
    'Albaranes Proveedor
    Me.Toolbar1.Buttons(16).visible = ComprobarBotonMenuVisible(Me.mnComAlbMan, Activado)
    Me.Toolbar1.Buttons(16).Enabled = Activado
    
    'Facturas Proveedor
    Me.Toolbar1.Buttons(17).visible = ComprobarBotonMenuVisible(Me.mnComHcoFacturas, Activado)
    Me.Toolbar1.Buttons(17).Enabled = Activado
    
    'Recepcion facturas de compras
    Me.Toolbar1.Buttons(18).visible = ComprobarBotonMenuVisible(Me.mnComFacturar, Activado)
    Me.Toolbar1.Buttons(18).Enabled = Activado


    '-----------------------------------------------------------
    'Mantenimientos
    Me.Toolbar1.Buttons(21).visible = ComprobarBotonMenuVisible(Me.mnManEntrada, Activado)
    Me.Toolbar1.Buttons(21).Enabled = Activado
    
    'Nº Serie
    Me.Toolbar1.Buttons(22).visible = ComprobarBotonMenuVisible(Me.mnRepNumSerie, Activado)
    Me.Toolbar1.Buttons(22).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Conuslta de precio
    Me.Toolbar1.Buttons(24).visible = ComprobarBotonMenuVisible(Me.mnFacPedidos(8), Activado)
    Me.Toolbar1.Buttons(24).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Gastos tecnicos
    'Para EULER --> Reloj
    mnobra(2).Caption = "Partes de trabajo"
    If vParamAplic.NumeroInstalacion = 4 Then
        Me.Toolbar1.Buttons(24).visible = ComprobarBotonMenuVisible(Me.mnobra(4), Activado)
        Me.Toolbar1.Buttons(24).Enabled = Activado
        mnobra(2).Caption = "Mantenimiento tareas reloj"
    Else
        Me.Toolbar1.Buttons(24).visible = ComprobarBotonMenuVisible(Me.mnAdmGastosTec, Activado)
        Me.Toolbar1.Buttons(24).Enabled = Activado
    End If
    
    'Nuevos botones
    'TPV
    Me.Toolbar1.Buttons(26).visible = ComprobarBotonMenuVisible(mnTPVLinea(0), Activado)
    If Activado Then
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "spatpvg", "1", "1")
        If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
        If Val(CadenaDesdeOtroForm) = 0 Then Activado = False
    End If
    Me.Toolbar1.Buttons(26).Enabled = Activado
    
    'Cambio empresa
   ' Me.Toolbar1.Buttons(27).visible = ComprobarBotonMenuVisible(mnCambioEmpresa, Activado)
    'Cambiar empresa lo dejo desde Febrero 2013 SIEMPRE visibñe
    Me.Toolbar1.Buttons(27).Enabled = True
    Me.Toolbar1.Buttons(27).visible = True
    
    'Agenda
    If vParamAplic.Frecuencias Then
        Me.Toolbar1.Buttons(28).Image = 24 'FRECUENCIAS
        Me.Toolbar1.Buttons(28).visible = ComprobarBotonMenuVisible(mnFrecuencias, Activado)
        Me.Toolbar1.Buttons(28).Enabled = Activado
        Me.Toolbar1.Buttons(28).ToolTipText = "Frecuencias"
    Else
        Me.Toolbar1.Buttons(28).Image = 20 'Agenda
        Me.Toolbar1.Buttons(28).visible = ComprobarBotonMenuVisible(mnAgenda, Activado)
        Me.Toolbar1.Buttons(28).Enabled = Activado
        Me.Toolbar1.Buttons(28).ToolTipText = "Agenda"
    End If
    
    'Avisos
    If vParamAplic.Ariagro = "" Then
        Me.Toolbar1.Buttons(23).visible = ComprobarBotonMenuVisible(mnRepAvisos, Activado)
        Me.Toolbar1.Buttons(23).Enabled = Activado
    Else
        'partes de trabajo
        Me.Toolbar1.Buttons(23).visible = ComprobarBotonMenuVisible(mnTratamientos(4), Activado)
        Me.Toolbar1.Buttons(23).Enabled = Activado
    End If
End Sub




Private Function ComprobarBotonMenuVisible(objMenu As Menu, Activado As Boolean) As Boolean
'Comprueba si el icono de la barra se debe activar/desactivar o si se debe poner
'visible o invisible. Para ello comprueba si su correspondiente entrada de menu
'esta activada/desactiva o visible/invisible
'(se comprueba hasta q se encuentra el false o se llega al padre)
Dim nomMenu As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cad As String
Dim b As Boolean


    On Error GoTo EComprobar
    
    b = objMenu.visible
    Activado = objMenu.Enabled
    
    If b = False Then
        ComprobarBotonMenuVisible = b
    Else
    
        nomMenu = objMenu.Name
        
        Set RS = New ADODB.Recordset
        
        'Obtener el padre del menu
        SQL = "select padre from usuarios.appmenus where aplicacion='Ariges' and name=" & DBSet(nomMenu, "T")
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            cad = RS.Fields(0).Value
        End If
        RS.Close
        
        b = True
        While b And cad <> ""
                SQL = "Select name,padre from usuarios.appmenus where aplicacion='Ariges' and contador= " & cad
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    cad = RS!Padre
                    nomMenu = RS!Name
                End If
                RS.Close
                
                'comprobar si el padre esta bloqueado
                SQL = "Select count(*) from usuarios.appmenususuario where aplicacion='Ariges' and codusu=" & Val(Right(CStr(vUsu.codigo), 3))
                SQL = SQL & " and tag='" & nomMenu & "|'"
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS.Fields(0).Value > 0 Then
                    'Esta bloqueado el menu para el usuario
                    b = False
                    Activado = False
                End If
                RS.Close
                If cad = "0" Then cad = "" 'terminar si llegamos a la raiz
        Wend
        ComprobarBotonMenuVisible = b
        Set RS = Nothing
    End If
    
EComprobar:
    If Err.Number <> 0 Then Err.Clear
End Function



Private Sub AbrirListado2(KOpcion As Integer)
    Screen.MousePointer = vbHourglass
    frmListado2.Opcion = KOpcion
    frmListado2.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub AbrirListado3(KOpcion As Integer)
    Screen.MousePointer = vbHourglass
    frmListado3.Opcion = KOpcion
    frmListado3.Show vbModal
    Screen.MousePointer = vbDefault
End Sub






'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'
'
'   ARIDOC.  para los datos de ARIDOC reutilizare la conneion conta
'           con lo cual la cerrare y abrire tantas veces necesite
'


Private Sub HacerMenuARidoc(Opcion As Byte)
    
    If Conexion_Aridoc_(True) Then
        Select Case Opcion
        Case 0
            frmAridocConfig.Show vbModal
        End Select
    End If
    Conexion_Aridoc_ False
End Sub














