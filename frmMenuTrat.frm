VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMenuTrat 
   BackColor       =   &H00858585&
   Caption         =   "Ariges 4"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9750
   Icon            =   "frmMenuTrat.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageListB 
      Left            =   5640
      Top             =   1320
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
            Picture         =   "frmMenuTrat.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":7C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8688
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":909A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":9AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":A4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":AED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":B8E2
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
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":C2F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":D386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":E418
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":F4AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1053C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":11FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":13050
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":140E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":15174
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":16206
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":17298
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1832A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":193BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1A44E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1B4E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1C572
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1D604
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1E696
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":1F728
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
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   29
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
            Style           =   3
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
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gastos técnicos"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Venta TPV"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   6615
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmMenuTrat.frx":210BA
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9128
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
            TextSave        =   "11:20"
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
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":2467C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":26386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":2C62C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":2D03E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":2DA50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":30202
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":30ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":313B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":31C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3256A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":32F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":333D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":334E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":335FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3370C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":33A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":39648
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3A05A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3AA6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3AB7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3B590
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3BFA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3C9B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3CCCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3CFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3D43A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3D88C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3DCDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3E130
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3E582
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3E9D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3ECEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3EE48
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3F162
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3F47C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":3FD56
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":40630
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":4094A
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":40AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":40DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":417D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":421E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":42BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":43606
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":44018
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":44A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":463BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":47D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":496E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":4B072
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":4CA04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":4E396
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
            Picture         =   "frmMenuTrat.frx":4FD28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":5017A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":505CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":50A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":50E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":512C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":51714
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":51B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":51FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":58252
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":58C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":5EEFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":65760
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":6BFC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":72824
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":79086
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":7F8E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8614A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8659C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":869EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":86E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":87292
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":876E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":87B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8D758
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8E5AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8E8C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8EBDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuTrat.frx":8EEF8
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
      Begin VB.Menu mnBarra1 
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
         Begin VB.Menu mnAlmListInactivos 
            Caption         =   "Listado Articulos &Inactivos"
         End
         Begin VB.Menu mnAlmListValoracion 
            Caption         =   "Listado Valoración &Stocks"
         End
         Begin VB.Menu mnAlmListMaxMin 
            Caption         =   "Inf. Stocks Máximos-Mínimos"
         End
         Begin VB.Menu mnAlmStockFecha 
            Caption         =   "Inf. Stocks a una &Fecha"
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
            Caption         =   "Actualizar &direrencias"
         End
         Begin VB.Menu mnAlmValoracionInve 
            Caption         =   "&Valoración stocks inventariados"
         End
         Begin VB.Menu mnBarra2 
            Caption         =   "-"
         End
         Begin VB.Menu mnAlmHcoInven 
            Caption         =   "&Histórico inventario"
         End
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
         Begin VB.Menu mnFacCartas 
            Caption         =   "Tipos de C&artas"
         End
         Begin VB.Menu mnFacIncidencias 
            Caption         =   "&Incidencias"
         End
      End
      Begin VB.Menu mnFacInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnFacInactivos 
            Caption         =   "Clientes Inacti&vos"
         End
         Begin VB.Menu mnFacInfClientes 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu mnFacAltas 
            Caption         =   "&Altas Clientes"
         End
         Begin VB.Menu mnFacEtiqClien 
            Caption         =   "&Etiquetas de clientes"
         End
         Begin VB.Menu mnFacCartaClien 
            Caption         =   "Car&tas a clientes"
         End
         Begin VB.Menu mnEtiquetasBultos 
            Caption         =   "&Etiquetas de bultos"
         End
      End
      Begin VB.Menu mnFacPreciosDtos 
         Caption         =   "&Precios y Descuentos"
         Begin VB.Menu mnFacTarifasVen 
            Caption         =   "&Tarifas Venta"
         End
         Begin VB.Menu mnFacTarifasPrecios 
            Caption         =   "&Lista Precios"
         End
         Begin VB.Menu mnFacPreEspecial 
            Caption         =   "Precios &Especiales"
         End
         Begin VB.Menu mnFacPromociones 
            Caption         =   "&Promociones"
         End
         Begin VB.Menu mnFacDescuentos 
            Caption         =   "&Descuentos Familia/Marca"
         End
         Begin VB.Menu mnFacBoniFact 
            Caption         =   "&Bonificaciones Factura"
         End
         Begin VB.Menu Barra12 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacInfMargenes 
            Caption         =   "&Control margenes tarifas"
         End
         Begin VB.Menu mnFactActPrecios 
            Caption         =   "&Actualizar precios"
         End
      End
      Begin VB.Menu mnFacOfertas 
         Caption         =   "&Ofertas"
         Begin VB.Menu mnFacEntOfertas 
            Caption         =   "&Mantenimiento Ofertas"
         End
         Begin VB.Menu mnFacGrupoPlant 
            Caption         =   "&Grupo de Plantillas"
         End
         Begin VB.Menu mnFacPlantillas 
            Caption         =   "Entrada de  &Plantillas"
         End
         Begin VB.Menu mnFacOfeEfectuadas 
            Caption         =   "Ofertas E&fectuadas"
         End
         Begin VB.Menu mnBarra3 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacHisOfertas 
            Caption         =   "&Histórico  Ofertas"
         End
         Begin VB.Menu mnFacTrasHist 
            Caption         =   "&Traspaso a Histórico"
         End
      End
      Begin VB.Menu mnFacPedidos 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnFacEntPedidos 
            Caption         =   "&Mantenimiento Pedidos"
         End
         Begin VB.Menu mnFacHcoPedidos 
            Caption         =   "&Histórico Pedidos Anulados"
         End
         Begin VB.Menu mnBarra4 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacConfirmPed 
            Caption         =   "&Cartas Confirmacion de Pedidos"
         End
         Begin VB.Menu mnFacPedidoxArtic 
            Caption         =   "Informe &Pedidos por Articulo"
         End
         Begin VB.Menu mnFacPedidoxClien 
            Caption         =   "Informe P&edidos por Cliente"
         End
         Begin VB.Menu mnFacDispStock 
            Caption         =   "Informe &Disponibilidad Stocks"
         End
      End
      Begin VB.Menu mnFacAlbaran 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnFacEntAlbaran 
            Caption         =   "&Mantenimiento Albaranes"
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
         Begin VB.Menu mnBarra5 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacPreFacturar 
            Caption         =   "&Previsión Facturación"
         End
         Begin VB.Menu mnFacFacturarAlb 
            Caption         =   "&Facturación de Albaranes"
         End
         Begin VB.Menu mnFacAlbMostrador 
            Caption         =   "Facturas de Mo&strador"
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
         Begin VB.Menu mnBarra9 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacContFactu 
            Caption         =   "&Contabilizar Facturas"
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
            Caption         =   "Ventas por &familia"
         End
         Begin VB.Menu mnFacEstDetalleFac 
            Caption         =   "&Detalle facturación"
         End
         Begin VB.Menu mnFacEstMargenVtas 
            Caption         =   "Mar&gen ventas por artículo "
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
         Begin VB.Menu mnComInfProve 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnComEtiqProve 
            Caption         =   "&Etiquetas de proveedores"
         End
         Begin VB.Menu mnComCartaProve 
            Caption         =   "&Cartas a Proveedores"
         End
      End
      Begin VB.Menu mnComPreciosDtos 
         Caption         =   "Precios y &Descuentos"
         Begin VB.Menu mnComPreProve 
            Caption         =   "P&recios Proveedor"
         End
         Begin VB.Menu mnComDtosProve 
            Caption         =   "Descuentos Pro&veedor"
         End
      End
      Begin VB.Menu mnComPedidos 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnComPedMant 
            Caption         =   "Mant. &Pedidos Proveedor"
         End
         Begin VB.Menu mnComHcoPedidos 
            Caption         =   "&Histórico Pedidos Anulados"
         End
         Begin VB.Menu mnComPteRecibir 
            Caption         =   "List. &Material pendiente de recibir"
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
      End
      Begin VB.Menu Barra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnComEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnComEstComprasxProve 
            Caption         =   "Compras por &Proveedor"
         End
         Begin VB.Menu mnComEstComprasxFam 
            Caption         =   "Compras por &Familia/Artíc."
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
      Begin VB.Menu mnBarra13 
         Caption         =   "-"
      End
      Begin VB.Menu mnManPrevFac 
         Caption         =   "&Previsión Facturación"
      End
      Begin VB.Menu mnManFactAlb 
         Caption         =   "Fac&turación  Mantenimientos"
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
      Begin VB.Menu mnRepMotivosPend 
         Caption         =   "Motivos &Pend. Rep."
      End
      Begin VB.Menu mnRepHistorico 
         Caption         =   "&Histórico de Reparaciones"
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
   End
   Begin VB.Menu mnTPV 
      Caption         =   "&Punto de Venta"
      Begin VB.Menu mnTPVpantallaVenta 
         Caption         =   "Pantalla de &venta"
      End
      Begin VB.Menu mnTPVcierreCaja 
         Caption         =   "&Cierre de caja"
      End
      Begin VB.Menu mnEtiqEstanteria 
         Caption         =   "Etiquetas estantería"
      End
      Begin VB.Menu mnBarra16 
         Caption         =   "-"
      End
      Begin VB.Menu mnTPVParamGen 
         Caption         =   "&Parámetros generales TPV"
      End
      Begin VB.Menu mnTPVParamTer 
         Caption         =   "Parámetros &terminales TPV"
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnBackUp 
         Caption         =   "&Copia Seguridad local"
      End
      Begin VB.Menu mnRecupFac 
         Caption         =   "&Recuperar facturas"
         Visible         =   0   'False
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
         Begin VB.Menu mnUtiMensNuevo 
            Caption         =   "&Nuevo"
         End
         Begin VB.Menu mnUtiMenEnvRec 
            Caption         =   "&Enviar/Recibir"
         End
         Begin VB.Menu mnBarra22 
            Caption         =   "-"
         End
         Begin VB.Menu mnUtiMensTipMen 
            Caption         =   "&Tipo de mensaje"
         End
      End
   End
   Begin VB.Menu mnSoporte 
      Caption         =   "&Soporte"
      Begin VB.Menu mnAyuda 
         Caption         =   "Ayuda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnEnviarMail 
         Caption         =   "Enviar Mail"
         Visible         =   0   'False
      End
      Begin VB.Menu mnWeb 
         Caption         =   "Web Ariadna Software"
      End
      Begin VB.Menu mnCheckVersion 
         Caption         =   "Comprobar version operativa"
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra11 
         Caption         =   "-"
      End
      Begin VB.Menu mnAcercaDe 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnFrecuencias 
      Caption         =   "Frecuencias"
   End
End
Attribute VB_Name = "frmMenuTrat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrimeraVez As Boolean

Dim TieneEditorDeMenus As Boolean


Private Sub MDIForm_Activate()
'Formulario Principal
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
    End If
    If Not vParam Is Nothing Then
        If vParam.Modificado Then
          'Poner datos visible del form
           PonerDatosVisiblesForm
           vParam.Modificado = False
        End If
    End If
    '--
    Screen.MousePointer = vbDefault
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

        .Buttons(15).Image = 9   'Pedidos Proveedor
        .Buttons(16).Image = 10   'Albaranes Proveedor
        .Buttons(17).Image = 11   'Facturas Proveedor
        .Buttons(18).Image = 12   'Recepcion Facturas Proveedor
        
        .Buttons(21).Image = 15   'Mantenimientos
        .Buttons(22).Image = 16   'Nº Serie

        .Buttons(24).Image = 13 'Gastos tecnicos
        .Buttons(26).Image = 19 'Pantalla venta del TPV
        
        .Buttons(29).Image = 14 'Salir
    End With
    LeerEditorMenus
    
    
    PonerDatosFormulario
    
       
'    Me.mnPedirPwd.Checked = vConfig.PedirPasswd
       
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\arifon2.dat")
    If Err.Number <> 0 Then
        Me.Picture = LoadPicture()
        Err.Clear
    End If
End Sub




Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vEmpresa Is Nothing) Or (vParam Is Nothing) Or (vParamAplic Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True

    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If Config Then HabilitarSoloPrametros_o_Empresas False
    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    
    'Comprobar que los iconos de la barra su correspondiente
    'entrada de menu esta habilitada sino desabilitar
    PoneBarraMenus
    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'Formulario Principal
Dim cad As String

    'Alguna cosilla antes de cerrar. Eliminar bloqueos
    cad = "Delete from zbloqueos where codusu = " & vUsu.Codigo
    Conn.Execute cad

    'Elimnar bloquo BD
    Set vUsu = Nothing
    Set vConfig = Nothing
    Set vEmpresa = Nothing
    
    Set vParam = Nothing
    Set vParamAplic = Nothing
    
    
    TerminaBloquear
    
    'cerrar las conexiones
    Conn.Close
    CerrarConexionConta

End Sub


Private Sub mnAcercaDe_Click()
    Screen.MousePointer = vbDefault
    frmMensajes.OpcionMensaje = 3
    frmMensajes.Show vbModal
End Sub

Private Sub mnAdmGastosTec_Click()
'Gastos Técnicos
    frmAdmGasTec.Show vbModal
End Sub

Private Sub mnAdmTrabajadores_Click()
    frmAdmTrabajadores.Show vbModal
End Sub

Private Sub mnAlmActualizarInve_Click()
    AbrirListado (14)
End Sub

Private Sub mnAlmAlPropios_Click()
    frmAlmAlPropios.Show vbModal
End Sub

Private Sub mnAlmArticulos_Click()
    frmAlmArticulos.Show vbModal
End Sub


Private Sub mnAlmCategoria_Click()
'categorias de articulos
    frmAlmCategorias.Show vbModal
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

Private Sub mnAlmStockFecha_Click()
'Informe de Stocks a una Fecha
    AbrirListado (19)
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

Private Sub mnBackUp_Click()
'Copia de seguridad de toda la base de datos
    frmBackUP.Show vbModal
End Sub

Private Sub mnCambioEmpresa_Click()
    Dim AntUSU As Usuario

    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If

    'Borramos temporal
    Conn.Execute "Delete from zbloqueos where codusu = " & vUsu.Codigo


    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"
    
'    Set AntUSU = vUsu
'    Set vUsu = Nothing
    frmLogin.Show vbModal
'    If vUsu Is Nothing Then
'        Set vUsu = AntUSU
'        Set AntUSU = Nothing
'        Exit Sub
'    End If

    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    Conn.Close
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
    If AbrirConexionConta() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
        End
    End If

    
    
    Set vEmpresa = Nothing
    'LeerEmpresaParametros
    
     'Carga los Datos Básicos de la empresa
    LeerDatosEmpresa
    
    
    'Carga los Niveles de cuentas de Contabilidad de la empresa
    LeerNivelesEmpresa
    
'    PonerDatosFormulario
    PonerDatosVisiblesForm

    'Ponemos primera vez a false
    PrimeraVez = True
    Me.SetFocus

    Screen.MousePointer = vbDefault
End Sub


Private Sub mnCheckVersion_Click()
'    Screen.MousePointer = vbHourglass
'    LanzaHome "webversion"
'    espera 2
'    Screen.MousePointer = vbDefault
End Sub


Private Sub mnComAlbMan_Click()
'Mantenimiento de Albaranes a Proveedor
    frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmComEntAlbaranes.EsHistorico = False
    frmComEntAlbaranes.Show vbModal
End Sub

Private Sub mnComCartaProve_Click()
'Cartas a proveedores
     AbrirListadoOfer (306) '306: Informe Cartas a Proveedores
End Sub

Private Sub mnComContFactu_Click()
'Contabilizar Facturas
    AbrirListado (224) 'Para pedir datos
End Sub

Private Sub mnComDirecciones_Click()
    frmComDirecciones.Show vbModal
End Sub

Private Sub mnComDtosProve_Click()
    frmComDtosFamMarca.Show vbModal
End Sub

Private Sub mnComEstComprasxFam_Click()
'Listado de compras por Familia
    AbrirListadoOfer (311)
End Sub

Private Sub mnComEstComprasxProve_Click()
'Listado de compras por proveedor
    AbrirListadoOfer (310)
End Sub

Private Sub mnComEtiqProve_Click()
'Etiquetas de proveedores
     AbrirListadoOfer (305) '305: Informe Etiquetas de Proveedores
End Sub

Private Sub mnComFacturar_Click()
    frmComFacturar.Show vbModal
End Sub

Private Sub mnComHcoAlbaranes_Click()
'Historico albaranes de compras a proveedores
    frmComEntAlbaranes.EsHistorico = True
    frmComEntAlbaranes.Show vbModal
End Sub

Private Sub mnComHcoFacturas_Click()
    frmComHcoFacturas.hcoCodMovim = ""
    frmComHcoFacturas.Show vbModal
End Sub

Private Sub mnComHcoPedidos_Click()
    frmComEntPedidos.EsHistorico = True
    frmComEntPedidos.Show vbModal
End Sub

Private Sub mnComInfProve_Click()
'Informe de Proveedores
    AbrirListado (58)   ': Informe Proveedores
End Sub

Private Sub mnComPedMant_Click()
'Mnatenimiento de Pedidos de compras
    frmComEntPedidos.EsHistorico = False
    frmComEntPedidos.Show vbModal
End Sub

Private Sub mnComPreProve_Click()
    frmComPreciosProv.Show vbModal
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

Private Sub mnComPteRecibir_Click()
'Listado de material pendiente de recibir
    AbrirListadoOfer (307) '307: List. Materia pte recibir
End Sub

Private Sub mnConfManteUsuarios_Click()
'Mantenimiento de Usuarios

      frmMantenusu.Show vbModal
End Sub

Private Sub mnConfParamAplic_Click()
'Parametros de la Aplicación
    frmConfParamAplic.Show vbModal
End Sub

Private Sub mnConfParamContables_Click()
'Parametros contables de la Aplicacion
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


Private Sub mnEtiqEstanteria_Click()
    AbrirListado 94
End Sub

Private Sub mnEtiquetasBultos_Click()
'Listado de etiquetas de los bultos
    AbrirListado 95
End Sub

Private Sub mnFacActividades_Click()
    frmFacActividades.Show vbModal
End Sub

Private Sub mnFacAgentesCom_Click()
    frmFacAgentesCom.Show vbModal
End Sub

Private Sub mnFacAlbMostrador_Click()
'Abre el formulario de Albaranes para introducir el Albaran de Mostrador
'y desde este generar la Factura de mostrador
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALM"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub


Private Sub mnFacAlbRectifica_Click()
'Facturas Rectificativas
    'Abre el formulario de Albaranes para introducir el Albaran Rectificativo
    'y desde este generar la Factura Rectificativa
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ART"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacAlbxArtic_Click()
'Informe de Albaranes por Articulo
    AbrirListadoPed (49)
End Sub

Private Sub mnFacAltas_Click()
'Informe de Altas de Nuevos Clientes
    AbrirListadoOfer (48) '48: Informes Altas Clientes
End Sub

Private Sub mnFacBancosPropios_Click()
    frmFacBancosPropios.Show vbModal
End Sub

Private Sub mnFacBoniFact_Click()
'Bonificacines factura
    frmFacBonificacion.Show vbModal
End Sub

Private Sub mnFacCartaClien_Click()
'Cartas a clientes
     AbrirListadoOfer (91) '91: Informe Cartas a Clientes
End Sub

Private Sub mnFacCartas_Click()
'Mantenimiento de Cartas
    frmFacCartasOferta.Show vbModal
End Sub


Private Sub mnFacClientes_Click()
'Mantenimiento de Clientes
    frmFacClientes.Show vbModal
End Sub

Private Sub mnFacClientesV1_Click()
'Mantenimiento de Clientes Varios
    frmFacClientesV.Show vbModal
End Sub

Private Sub mnFacConfirmPed_Click()
    AbrirListadoOfer (40)
End Sub

Private Sub mnFacContFactu_Click()
'Contabilizar Facturas
    AbrirListado (223) 'Para pedir datos
End Sub

Private Sub mnFacDescuentos_Click()
    frmFacDtosFamMarca.Show vbModal
End Sub

Private Sub mnFacDispStock_Click()
'Resumen de Disponibilidad de Stocks
    AbrirListadoPed (42)
End Sub

Private Sub mnFacEntAlbaran_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALV"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacEntOfertas_Click()
    frmFacEntOfertas.EsHistorico = False
    frmFacEntOfertas.Show vbModal
End Sub

Private Sub mnFacEntPedidos_Click()
'Mantenimiento de Pedidos
    frmFacEntPedidos.EsHistorico = False
    frmFacEntPedidos.Show vbModal
End Sub

Private Sub mnFacEstDetalleFac_Click()
'Detalle facturacion clientes
     AbrirListadoOfer (231)
End Sub

Private Sub mnFacEstMargenVtas_Click()
'Estadistica margen ventas por artículo
    AbrirListado (246)
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

Private Sub mnFacEtiqClien_Click()
'Etiquetas de clientes
     AbrirListadoOfer (90) '90: Informe Etiquetas de Clientes
End Sub

Private Sub mnFacFacturarAlb_Click()
'Facturacion de Albaranes de Ventas
    frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (52)
End Sub

Private Sub mnFacFormasPago_Click()
    frmFacFormasPago.Show vbModal
End Sub


Private Sub mnFacGrupoPlant_Click()
'Mantenimiento de Grupos de Plantillas
    frmFacGrupoPlantilla.Show vbModal
End Sub

Private Sub mnFacHcoAlbaranes_Click()
'Histórico de Albaranes eliminados
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALV"
    frmFacEntAlbaranes.EsHistorico = True
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacHcoFacturas_Click()
'Histórico de Facturas
    frmFacHcoFacturas.hcoCodMovim = ""
    frmFacHcoFacturas.Show vbModal
End Sub

Private Sub mnFacHcoPedidos_Click()
'Histórico de Pedidos
    frmFacEntPedidos.EsHistorico = True
    frmFacEntPedidos.Show vbModal
End Sub

Private Sub mnFacHisOfertas_Click()
'Histórico de Ofertas
    frmFacEntOfertas.EsHistorico = True
    frmFacEntOfertas.Show vbModal
End Sub

Private Sub mnFacInactivos_Click()
'Informe de Clientes Inactivos
    AbrirListadoOfer (46) '46: Informes Clientes Inactivos
End Sub

Private Sub mnFacIncidencias_Click()
    frmIncidencias.Show vbModal
End Sub

Private Sub mnFacIncumPlazos_Click()
'Incumplimiento de los Plazos de Entrega
    
    AbrirListadoPed (51)
End Sub

Private Sub mnFacInfClientes_Click()
'Informe de Clientes
    AbrirListadoOfer (47) '47: Informes Clientes
End Sub

Private Sub mnFacInfMargenes_Click()
'Informe control margenes de tarifas
    AbrirListado (245)
End Sub

Private Sub mnFacOfeEfectuadas_Click()
'Listado de Ofertas Efectuadas
    AbrirListadoOfer (34) '34: Informe Ofertas Efectuadas
End Sub

Private Sub mnFacPedidoxArtic_Click()
'Informe de Pedidos por Articulo
    AbrirListadoPed (41)
End Sub

Private Sub mnFacPedidoxClien_Click()
'Informe de Pedidos por Cliente
    AbrirListadoPed (44)
End Sub

Private Sub mnFacPlantillas_Click()
'Mantenimiento de Plantillas
    frmFacPlantilla.Show vbModal
End Sub

Private Sub mnFacPreEspecial_Click()
    frmFacPreciosEspecial.Show vbModal
End Sub

Private Sub mnFacPreFacturar_Click()
' Previsión Facturacion de Albaranes
    frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO
End Sub

Private Sub mnFacPromociones_Click()
    frmFacPromociones.Show vbModal
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

Private Sub mnFactActPrecios_Click()
'Actualizar precios actuales y especiales
    frmFacActPrecios.Show vbModal
End Sub

Private Sub mnFacTarifasPrecios_Click()
'Listado Precios
    frmFacTarifasPrecios.Show vbModal
End Sub

Private Sub mnFacTarifasVen_Click()
'Tarifas Venta
     frmFacTarifas.Show vbModal
End Sub


Private Sub mnFacTrasHist_Click()
'Traspaso de Ofertas a las tablas de Historico
    frmListadoOfer.OpcionListado = 36
    AbrirListadoOfer (36) 'NO IMPRIME LISTADO, hace traspaso de Ofertas de la tabla (scapre) a (schpre)
End Sub

Private Sub mnFacZonas_Click()
    frmFacZonas.Show vbModal
End Sub

Private Sub mnFacFormasEnvio_Click()
    frmFacFormasEnvio.Show vbModal
End Sub

Private Sub mnFrecuencias_Click()
    frmFrecuencias.Show vbModal
End Sub

Private Sub mnManAltas_Click()
'Listado Altas de Mantenimientos
    AbrirListado 73
End Sub

Private Sub mnManEntrada_Click()
    frmManMantenimientos.Show vbModal
End Sub

Private Sub mnManFactAlb_Click()
'Facturacion de Mantenimientos
     AbrirListadoPed (75) 'NO IMPRIME LISTADO
End Sub

Private Sub mnManFichas_Click()
'Listado Fichas de Mantenimientos
    AbrirListado 72
End Sub

Private Sub mnManListado_Click()
'Listados de Mantenimientos
    AbrirListado 70
End Sub

Private Sub mnManPrevFac_Click()
' Previsión Facturacion de Albaranes de Mantenimiento
'    frmListadoPed.CodClien = "ALM" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (74) 'NO IMPRIME LISTADO
End Sub

Private Sub mnManRevisiones_Click()
'Listado Revisiones de Mantenimientos
     AbrirListado 71
End Sub

Private Sub mnManTiposContrato_Click()
    frmManTiposContrato.Show vbModal
End Sub


Private Sub mnRecupFac_Click()
'recuperar facturas
    'abrimos albaranes de mostrador
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALM"
    'le indicamos q estamos recuperando facturas
    frmFacEntAlbaranes.RecuperarFactu = True
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnRepAlbaranes_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALR"
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnRepAvisos_Click()
'Avisos de averias de clientes
    frmRepAvisos.Show vbModal
End Sub

Private Sub mnRepControlRep_Click()
'Control de Reparaciones (para los Tecnicos)
    frmRepEntReparaciones.ControlRep = True
    frmRepEntReparaciones.EsHistorico = False
    frmRepEntReparaciones.Show vbModal
End Sub

Private Sub mnRepEntReparacion_Click()
'Mantenimiento de Reparaciones
    frmRepEntReparaciones.ControlRep = False
    frmRepEntReparaciones.EsHistorico = False
    frmRepEntReparaciones.Show vbModal
End Sub

Private Sub mnRepFactAlb_Click()
'Facturacion de Albaranes de Reparacion
    frmListadoPed.CodClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (52)
End Sub

Private Sub mnRepHistorico_Click()
'Historico de las reparaciones
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

Private Sub mnRepMotivosPend_Click()
'Motivos Pendientes Reparar
    frmRepMotivosPend.Show vbModal
End Sub

Private Sub mnRepNumSerie_Click()
'Mantenimiento de Nºs de Serie
    frmRepNumSerie.Show vbModal
End Sub

Private Sub mnRepPrevFact_Click()
' Previsión Facturacion de Albaranes de Reparacion
    frmListadoPed.CodClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO

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

Private Sub mnTiposArticulos_Click()
    frmAlmTipoArticulo.Show vbModal
End Sub

Private Sub mnSalir_Click()
    End
End Sub





Private Sub mnTPVcierreCaja_Click()
'Abre el informe de cierre de caja del dia en el TPV
    AbrirListadoOfer (240)
End Sub

Private Sub mnTPVpantallaVenta_Click()
'Pantalla venta del TPV
Dim Nom As String

    'Antes de abrir la pantalla de venta comprobamos que podemos leer el terminal
    'nom = ComputerNameTServer

    Nom = ComputerName 'Nombre PC conectado por Terminal Server / local
    
    If Trim(Nom) <> "" Then
        frmFacTPVEnt.NomrePC_conectado = Nom
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


Private Sub mnTPVParamGen_Click()
'Parámetros generales del TPV
    frmFacTPVParamG.Show vbModal
End Sub

Private Sub mnTPVParamTer_Click()
'parametros de los terminales(equipos) del TPV
    frmFacTPVParamT.Show vbModal
End Sub

Private Sub mnUtiBuscarErrConCli_Click()
'Facturas pendientes de contabilizar (CLIENTES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 6
    frmUtilidades.Show vbModal
End Sub

Private Sub mnUtiBuscarErrConPro_Click()
'Facturas pendientes de contabilizar (PROVEEDORES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 7
    frmUtilidades.Show vbModal
End Sub


Private Sub mnUtiBuscarErrFac_Click()
'Buscar errores en nº de factura (solo en facturas de clientes)
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 5
    frmUtilidades.Show vbModal
End Sub



Private Sub mnUtiConnActivas_Click()
'ver las conexiones a donde apuntan
Dim cad As String
    cad = "Conexiones:" & vbCrLf
    cad = cad & "------------------" & vbCrLf & vbCrLf
    cad = cad & "Ariges: " & vbCrLf & Conn.ConnectionString & vbCrLf & vbCrLf
    cad = cad & "Conta: " & vbCrLf & ConnConta.ConnectionString & vbCrLf
    MsgBox cad, vbInformation
End Sub

Private Sub mnUtiMensNuevo_Click()
'Nuevo mensaje en la utilidad de mensajeria interna
    frmMensaje2.Show vbModal
End Sub

Private Sub mnUtiMensTipMen_Click()
    frmTiposMensajes.Show vbModal
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



Private Sub mnWeb_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome ("websoporte")
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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
        mnFacEntOfertas_Click
    Case 10 'Pedidos a Clientes
        mnFacEntPedidos_Click
    Case 11 'Albaranes a Clientes
        mnFacEntAlbaran_Click
    Case 12 'Hist. Albaranes (Facturas)
        mnFacHcoFacturas_Click
        
    Case 15 'Pedidos de Proveedores
        mnComPedMant_Click
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
        
    Case 24 'Gastos Técnicos
        mnAdmGastosTec_Click
        
    Case 26 'Entrada al TPV
        mnTPVpantallaVenta_Click
        
    Case 29 'Salir
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
            If LCase(Mid(T.Name, 1, 6)) <> "mnbarr" Then _
                T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnConfParamAplic = True
    Me.mnConfParamGenerales = True

    Me.mnSalir.Enabled = True
    Me.mnCambioEmpresa.Enabled = True
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
        Me.mnFacTrasHist.Enabled = False
        
        'Facturacion Ventas
        Me.mnFacFacturarAlb.Enabled = False
        Me.mnFacContFactu.Enabled = False
        
        'Facturacion Compras
        Me.mnComFacturar.Enabled = False
        Me.mnComContFactu.Enabled = False
        
        'Reparaciones
        Me.mnRepFactAlb.Enabled = False
        
        'Mantenimientos
        Me.mnManFactAlb.Enabled = False
    End If
End Sub



Private Sub LanzaHome(opcion As String)
Dim i As Integer
Dim cad As String

    On Error GoTo ELanzaHome

'    LanzaHome = False
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", opcion, "codigo", "1", "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en Parámetros de la Aplicación.", vbExclamation
        Exit Sub
    End If

    If opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision


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
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    
    SQL = "Select * from usuarios.appmenususuario where aplicacion='Ariges' and codusu = " & Val(Right(CStr(vUsu.Codigo), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3) & "·"
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
                If InStr(1, SQL, C) > 0 Then T.visible = False
           
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



Private Sub PoneBarraMenus()
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
    Me.Toolbar1.Buttons(9).visible = ComprobarBotonMenuVisible(Me.mnFacEntOfertas, Activado)
    Me.Toolbar1.Buttons(9).Enabled = Activado
    
    'Pedidos Clientes
    Me.Toolbar1.Buttons(10).visible = ComprobarBotonMenuVisible(Me.mnFacEntPedidos, Activado)
    Me.Toolbar1.Buttons(10).Enabled = Activado
    
    'Albaranes Clientes
    Me.Toolbar1.Buttons(11).visible = ComprobarBotonMenuVisible(Me.mnFacEntAlbaran, Activado)
    Me.Toolbar1.Buttons(11).Enabled = Activado
    
    'Facturas Clientes
    Me.Toolbar1.Buttons(12).visible = ComprobarBotonMenuVisible(Me.mnFacHcoFacturas, Activado)
    Me.Toolbar1.Buttons(12).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Pedidos Proveedor
    'Comprobar que los menus del que cuelga no este bloqueado o invisible
        Me.Toolbar1.Buttons(15).visible = ComprobarBotonMenuVisible(Me.mnComPedMant, Activado)
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
    'Gastos tecnicos
    Me.Toolbar1.Buttons(25).visible = ComprobarBotonMenuVisible(Me.mnAdmGastosTec, Activado)
    Me.Toolbar1.Buttons(25).Enabled = Activado
    
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
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            cad = RS.Fields(0).Value
        End If
        RS.Close
        
        b = True
        While b And cad <> ""
                SQL = "Select name,padre from usuarios.appmenus where aplicacion='Ariges' and contador= " & cad
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    cad = RS!Padre
                    nomMenu = RS!Name
                End If
                RS.Close
                
                'comprobar si el padre esta bloqueado
                SQL = "Select count(*) from usuarios.appmenususuario where aplicacion='Ariges' and codusu=" & vUsu.Codigo
                SQL = SQL & " and tag='" & nomMenu & "|'"
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS.Fields(0).Value > 0 Then
                    'Esta bloqueado el menu para el usuario
                    b = False
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
