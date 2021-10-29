VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.ShortcutBar.v17.2.0.ocx"
Begin VB.Form frmShortcutBar2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeShortcutBar.ShortcutBar wndShortcutBar 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
      _Version        =   1114114
      _ExtentX        =   4683
      _ExtentY        =   11033
      _StockProps     =   64
      VisualTheme     =   3
      MinimumClientHeight=   20
      AllowMinimize   =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ariadna "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   120
      Picture         =   "frmShortcutBar.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmShortcutBar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShortcutBarMinimized As Boolean

Private Sub Form_Load()
  
    CreateShortcutBar
    
    wndShortcutBar.FindItem(SHORTCUT_CONTACTS).Selected = True
    
    wndShortcutBar.ExpandedLinesCount = 6
    wndShortcutBar.EnableAnimation = True
    wndShortcutBar.ShowExpandButton = False
    ShortcutBarMinimized = False
    
End Sub

Sub CreateShortcutBar()

    Dim Item As ShortcutBarItem
    Set frmPaneCalendar2 = Nothing
    Set frmPaneContacts2 = Nothing
    Set frmPaneCalendar = New frmPaneCalendar2
    Set frmPaneContacts = New frmPaneContacts2
    Load frmPaneCalendar
    Load frmPaneContacts
    'Load frmPaneInformacion
    'Load frmPaneAcercaDe
    
  
    Set Item = wndShortcutBar.AddItem(SHORTCUT_CONTACTS, "Empresas", frmPaneContacts.hwnd)
    Set Item = wndShortcutBar.AddItem(SHORTCUT_CALENDAR, "Calendario", frmPaneCalendar.hwnd)
    
    
   
    
End Sub

Private Sub Label3_Click()

End Sub

Private Sub wndShortcutBar_SelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)

    CambioPane Item.Id, True
End Sub
Public Sub CambioPane(ItemId As Long, DesdeEsteForm As Boolean)

Dim TabNuevo As RibbonTab
    Select Case ItemId
        Case SHORTCUT_CONTACTS:

            frmInbox.CalendarControl.visible = False
            frmInbox.ScrollBarCalendar.visible = False
            Set TabNuevo = frmppalN.RibbonBar.FindTab(3)  'Diario
            
           
        Case SHORTCUT_CALENDAR:
            frmPaneCalendar.AsignarCalendar
            frmInbox.CalendarControl.visible = True
            frmInbox.ScrollBarCalendar.visible = True
            Set TabNuevo = frmppalN.RibbonBar.FindTab(ID_TAB_CALENDAR_HOME)
            
            frmPaneCalendar.Enabled = True
            frmPaneCalendar.DatePicker.Enabled = True

    End Select
    If Not TabNuevo Is Nothing And DesdeEsteForm Then
    
        TabNuevo.Selected = True
        Set TabNuevo = Nothing
    End If
    
    frmInbox.Form_Resize
    
   ' Debug.Print Item.IconHandle
End Sub


Public Sub Form_Resize()
    On Error Resume Next
    
    Dim nWidth As Long
    Dim nHe As Long
    nWidth = Me.ScaleWidth - 8
    
    
    wndShortcutBar.Move 4, Image1.Height + 6, nWidth, ScaleHeight - 6 - Image1.Height
       
       
       
    Dim Minimized As Boolean
    Minimized = wndShortcutBar.Width <= wndShortcutBar.MinimizedWidth
    If (Minimized <> ShortcutBarMinimized) Then
        ShortcutBarMinimized = Minimized
        
        
        frmPaneCalendar.MainCaption.Expanded = Not Minimized
        frmPaneContacts.MainCaption.Expanded = Not Minimized
'        frmPaneInformacion.MainCaption.Expanded = Not Minimized
'        frmPaneAcercaDe.MainCaption.Expanded = Not Minimized
        'frmPaneFolders.MainCaption.Expanded = Not Minimized
        'frmPaneShortcuts.MainCaption.Expanded = Not Minimized
        'frmPaneJournal.MainCaption.Expanded = Not Minimized
    End If

End Sub


Public Sub SetColor(Id As Integer)
    Set wndShortcutBar.Icons = CommandBarsGlobalSettings.Icons
    Me.BackColor = wndShortcutBar.PaintManager.PaneBackgroundColor
   ' Me.Image1.Visible = vUsu.Skin = 2
    Me.Image1.visible = True 'vUsu.Skin <> 2
    
    
    If Id = ID_OPTIONS_STYLEBLACK2010 Then
        Label1.ForeColor = vbWhite
        Label2.ForeColor = &HE0E0E0
    ElseIf Id = ID_OPTIONS_STYLESILVER2010 Then
        'HexColor = &H73716B
        Label1.ForeColor = vbBlack
        Label2.ForeColor = vbWhite
    Else
        Label1.ForeColor = &H800000
        Label2.ForeColor = vbWhite
    End If
    
End Sub

