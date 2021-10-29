VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#17.2#0"; "Codejock.SkinFramework.v17.2.0.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.CommandBars.v17.2.0.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.DockingPane.v17.2.0.ocx"
Begin VB.Form frmppalN 
   Caption         =   "Ariges6"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16755
   FillStyle       =   0  'Solid
   Icon            =   "frmPpalN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   16755
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8880
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FC8A
            Key             =   "New"
            Object.Tag             =   "100"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FCE8
            Key             =   "Open"
            Object.Tag             =   "101"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FD46
            Key             =   "Save"
            Object.Tag             =   "103"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FDA4
            Key             =   "Print"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FE02
            Key             =   "Cut"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FE60
            Key             =   "Copy"
            Object.Tag             =   "106"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FEBE
            Key             =   "Paste"
            Object.Tag             =   "107"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FF1C
            Key             =   "Bold"
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FF7A
            Key             =   "Italic"
            Object.Tag             =   "121"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FFD8
            Key             =   "Underline"
            Object.Tag             =   "122"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":10036
            Key             =   "Align Left"
            Object.Tag             =   "123"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":10094
            Key             =   "Center"
            Object.Tag             =   "124"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":100F2
            Key             =   "Align Right"
            Object.Tag             =   "125"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":10150
            Key             =   "About"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":101AE
            Key             =   ""
            Object.Tag             =   "166"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1020C
            Key             =   ""
            Object.Tag             =   "168"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1026A
            Key             =   ""
            Object.Tag             =   "165"
         EndProperty
      EndProperty
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   9480
      Top             =   1560
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   6960
      Top             =   1800
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":102C8
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   6960
      Top             =   600
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   6120
      Top             =   960
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager ImageManagerGalleryStyles 
      Left            =   8040
      Top             =   240
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":102E2
   End
End
Attribute VB_Name = "frmppalN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Dim ContextEvent As CalendarEvent


Dim MRUShortcutBarWidth


Const IMAGEBASE = 10000
Const MinimizedShortcutBarWidth = 32 + 8

Dim WithEvents statusBar  As XtremeCommandBars.statusBar
Attribute statusBar.VB_VarHelpID = -1
Dim FontSizes(4) As Integer
Dim RibbonSeHaCreado As Boolean
Dim Pane As Pane
Dim Cad As String

'Variables comunes para todos los procedimientos de carga menus en el ribbon
'Codejock
Dim TabNuevo As RibbonTab
Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup
Dim GroupManageCalendars As RibbonGroup, GroupShare As RibbonGroup, GroupFind As RibbonGroup

Dim Control As CommandBarControl
Dim ControlNew_NewItems As CommandBarPopup
Dim Rn2 As ADODB.Recordset
Dim Habilitado As Boolean


Dim PrimeraVez As Boolean

Dim idMenuIconoAgrupados2 As Integer  'del 100 al 135

Dim T1 As Single






'Vamos a necesitar un array de arrays
'Cada bloque de cuatro son:
'   11|9|5|11|  Total: 40-> Necesito, de momento 10 bloques
 Dim MacroArray() As Variant





Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
End Function

Sub LoadResources(DllName As String, IniFileName As String)
Dim elpath As String
    
    elpath = App.Path & "\Styles\"
    CommandBarsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ShortcutBarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    SuiteControlsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    CalendarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ReportControlGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    DockingPaneGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
End Sub

Public Sub CheckButton(nButton As Integer)
    CommandBars.Actions(ID_OPTIONS_STYLEBLUE2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLESILVER2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLEBLACK2010).Checked = False
    
    CommandBars.Actions(nButton).Checked = True
End Sub

Sub OnThemeChanged(Id As Integer)
Dim N_Skin As Integer
    CheckButton Id
    
    Dim FlatStyle As Boolean
    FlatStyle = Id >= ID_OPTIONS_STYLESCENIC7 And Id <= ID_OPTIONS_STYLEBLACK2010
        
        
    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
   
    
    CommandBars.EnableOffice2007Frame False

    Select Case CommandBars.VisualTheme
        Case xtpThemeResource, xtpThemeRibbon
            CommandBars.AllowFrameTransparency False 'True
            CommandBars.EnableOffice2007Frame True
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
        Case Else
            CommandBars.AllowFrameTransparency True
            CommandBars.EnableOffice2007Frame False
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
    End Select
    
    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = CommandBars.ToolTipContext
    ToolTipContext.Style = xtpToolTipResource
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    
    statusBar.ToolTipContext.Style = ToolTipContext.Style
    frmShortBar.wndShortcutBar.ToolTipContext.Style = ToolTipContext.Style
    
       
    'CreateBackstage
    'SetBackstageTheme
    
    'CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\styles\Ariconta.ico", 16, 16
            
    'Set Captions VisualTheme
    On Error Resume Next
    Dim CtrlCaption As ShortcutCaption
    Dim Form As Form, Ctrl As Object
            
    For Each Form In Forms
       
        For Each Ctrl In Form.Controls
                    
            Set CtrlCaption = Ctrl
            If Not CtrlCaption Is Nothing Then
                CtrlCaption.VisualTheme = frmShortBar.wndShortcutBar.VisualTheme
            End If
                    
        Next
    Next
       
    DockingPaneManager.PaintManager.SplitterSize = 5
    DockingPaneManager.PaintManager.SplitterColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
    
    DockingPaneManager.PaintManager.ShowCaption = False
    DockingPaneManager.RedrawPanes
        
    frmShortBar.SetColor Id
    frmInbox.SetColor Id
        

    frmPaneCalendar.SetFlatStyle FlatStyle
    frmPaneContacts.SetFlatStyle FlatStyle
    'frmPaneInformacion.SetFlatStyle FlatStyle
    'frmPaneAcercaDe.SetFlatStyle FlatStyle
    
    
    
    
    
    
 
    LoadIcons
    N_Skin = Id - 2895
    EstablecerSkin N_Skin
    
    'Updatear SKIN usuario
    If CStr(N_Skin) <> vUsu.Skin Then
        vUsu.Skin = N_Skin
        vUsu.ActualizarSkin
    End If
    
End Sub

Public Sub SetBackstageTheme()
Dim i As Integer
    Dim nTheme As XtremeCommandBars.XTPBackstageButtonControlAppearanceStyle
    nTheme = xtpAppearanceResource

   ' If Not (pageBackstageInfo Is Nothing) Then
        'pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        'pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        'pageBackstageInfo.btnCheckForIssues.Appearance = nTheme
        'pageBackstageInfo.btnManageVersions.Appearance = nTheme
   ' End If
    
    If Not (pageBackstageHelp Is Nothing) Then
        For i = 0 To 4
            pageBackstageHelp.btnAcciones(i).Appearance = nTheme
        Next
        
    End If
    
    'If Not (pageBackstageSend Is Nothing) Then
        'pageBackstageSend.btnTab(0).Appearance = nTheme
        'pageBackstageSend.btnTab(1).Appearance = nTheme
        'pageBackstageSend.btnTab(2).Appearance = nTheme
        'pageBackstageSend.btnTab(3).Appearance = nTheme
    'End If

End Sub

Private Sub CreateStatusBar()
Dim Pane As StatusBarPane

    If RibbonSeHaCreado Then
        'StatusBar.Pane(0).Value = vEmpresa.nomempre & "    " & vUsu.Login
        statusBar.Pane(0).Text = "Nº " & vEmpresa.codempre
        statusBar.Pane(1).Text = vEmpresa.nomempre
    
    Else
    
         
         Set statusBar = Nothing
         
         Set statusBar = CommandBars.statusBar
         statusBar.visible = True
         
         
         Set Pane = statusBar.AddPane(ID_INDICATOR_PAGENUMBER)
         Pane.Text = "Nº " & vEmpresa.codempre
         Pane.Caption = "&C"
         Pane.Value = vEmpresa.nomempre & "    " & vUsu.Login
         Pane.Button = True
         Pane.SetPadding 8, 0, 8, 0
         
         Set Pane = statusBar.AddPane(ID_INDICATOR_WORDCOUNT)
         Pane.Text = vEmpresa.nomempre
         Pane.Caption = ""
         Pane.Value = vEmpresa.codempre
         Pane.Button = True
         Pane.SetPadding 8, 0, 8, 0
         
         
         Set Pane = statusBar.AddPane(0)
         Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
         Pane.BeginGroup = True
                 
        '
         statusBar.RibbonDividerIndex = 3
         statusBar.EnableCustomization True
         
         CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowNever
         CommandBars.Options.ShowKeyboardTips = True
         CommandBars.Options.ToolBarAccelTips = True
         
    End If
End Sub

Private Sub DockBarRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.RecalcLayout
    BarOnLeft.GetWindowRect Left, top, Right, Bottom
    
    CommandBars.DockToolBar BarToDock, Right, (Bottom + top) / 2, BarOnLeft.Position

End Sub

Private Sub CommandBars_CommandBarKeyDown(CommandBar As XtremeCommandBars.ICommandBar, KeyCode As Long, Shift As Integer)
    'Debug.Print CommandBar.BarID
End Sub

Public Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    HACER_CommandBars_Execute Control.Id
End Sub

Public Sub HACER_CommandBars_Execute(Control_Id As Long)
Dim AbiertoFormulario  As Boolean
    AbiertoFormulario = False
    
    
   ' Debug.Print Now & ": " & Control.Id
    Select Case Control_Id
        Case XTPCommandBarsSpecialCommands.XTP_ID_RIBBONCONTROLTAB:
            'If PrimeraVez Then S top
            
            
            If frmInbox.CalendarControl.visible = True Then
                If UCase(frmppalN.RibbonBar.SelectedTab.Caption) <> "AGENDA" Then
                   
                   frmShortcutBar2.CambioPane SHORTCUT_CONTACTS, False
                End If
            Else
                If UCase(frmppalN.RibbonBar.SelectedTab.Caption) = "AGENDA" Then
                   frmShortcutBar2.CambioPane SHORTCUT_CALENDAR, False
                End If
            End If
            
                
            
        Case XTP_ID_RIBBONCUSTOMIZE:
            CommandBars.ShowCustomizeDialog 3
            
        Case ID_APP_ABOUT:
            Dim DireccionAyuda As String
            
           LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "AriCONTA-6.html?"
   
        
        Case ID_FILE_NEW:
            'frmEmail.Show 0, Me
        
        
        'FALTA#
        'Case ID_Licencia_Usuario_Final_txt, ID_Licencia_Usuario_Final_web, ID_Ver_Version_operativa_web
        '    OpcionesMenuInformacion Control.Id
        
        
        
        Case ID_VIEW_STATUSBAR:
            CommandBars.statusBar.visible = Not CommandBars.statusBar.visible
            CommandBars.RecalcLayout
            
        Case ID_RIBBON_EXPAND:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
            Dim newFontHeight As Integer
            newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
            RibbonBar.FontHeight = newFontHeight
            
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
            CommandBars.PaintManager.AutoResizeIcons = Not CommandBars.PaintManager.AutoResizeIcons
            CommandBars.RecalcLayout
            RibbonBar.RedrawBar
            
        Case ID_OPTIONS_STYLEBLUE2010:
        
            LoadResources "Office2010.dll", "Office2010Blue.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLUE2010
            
            
            
       Case ID_OPTIONS_STYLESILVER2010:
            LoadResources "Office2010.dll", "Office2010Silver.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLESILVER2010
        
       Case ID_OPTIONS_STYLEBLACK2010:
            LoadResources "Office2010.dll", "Office2010Black.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLACK2010
        
        Case ID_APP_EXIT:
            Unload Me
        
    
            
        Case ID_GROUP_GOTO_TODAY:
            Select Case frmInbox.CalendarControl.ViewType
                Case xtpCalendarDayView:
                    frmInbox.CalendarControl.DayView.ShowDay DateTime.Now, True
            
                Case xtpCalendarWorkWeekView:
                    frmInbox.CalendarControl.DayView.SetSelection DateTime.Now, DateTime.Now, True
                    frmInbox.CalendarControl.RedrawControl
            
                Case xtpCalendarWeekView:
                    frmInbox.CalendarControl.WeekView.SetSelection DateTime.Now, DateTime.Now, True
            
                Case xtpCalendarMonthView:
                    frmInbox.CalendarControl.MonthView.SetSelection DateTime.Now, DateTime.Now, True
            End Select
            
        Case ID_GROUP_GOTO_NEXT7DAYS:
            Dim lastDate As Date
            lastDate = frmInbox.CalendarControl.DayView.Days(frmInbox.CalendarControl.DayView.DaysCount - 1).Date
            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
            frmInbox.CalendarControl.DayView.ShowDays lastDate + 1, lastDate + 7
            
        Case ID_GROUP_ARRANGE_DAY:
            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView
            
        Case ID_GROUP_ARRANGE_WEEK:
            frmInbox.CalendarControl.UseMultiColumnWeekMode = True
            frmInbox.CalendarControl.ViewType = xtpCalendarWeekView

        Case ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_MONTH_LOW, _
             ID_GROUP_ARRANGE_MONTH_MEDIUM, ID_GROUP_ARRANGE_MONTH_HIGH:
            frmInbox.CalendarControl.ViewType = xtpCalendarMonthView
            
        Case ID_CALENDAREVENT_OPEN:
            frmInbox.mnuOpenEvent
            
        Case ID_CALENDAREVENT_DELETE:
            frmInbox.mnuDeleteEvent
            
        Case ID_CALENDAREVENT_NEW, ID_GROUP_NEW_APPOINTMENT:
            'falta### frmEditEvent.AllDayOverride = False
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_MEETING:
            'falta### frmEditEvent.AllDayOverride = False
            'falta### frmEditEvent.chkMeeting.Value = 1
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_ALLDAY:
            'falta### frmEditEvent.AllDayOverride = True
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
        Case ID_GROUP_SHARE_SHARE
            frmInbox.CalendarControl.PrintOptions.Footer.TextCenter = vEmpresa.nomempre
            frmInbox.CalendarControl.PrintOptions.Footer.TextLeft = "Ariconta6. Ariadna SW"
            frmInbox.CalendarControl.PrintOptions.Footer.TextRight = Format(Now, "dd/mm/yyyy hh:mm")
            frmInbox.CalendarControl.PrintPreviewOptions.Title = "Ariconta6 " & vEmpresa.nomempre
            frmInbox.CalendarControl.PrintPreview True
            
        Case ID_CALENDAREVENT_CHANGE_TIMEZONE:
            frmInbox.mnuChangeTimeZone
            
        Case ID_CALENDAREVENT_60:
            frmInbox.mnuTimeScale 60
            
        Case ID_CALENDAREVENT_30:
            frmInbox.mnuTimeScale 30
            
        Case ID_CALENDAREVENT_15:
            frmInbox.mnuTimeScale 15
            
        Case ID_CALENDAREVENT_10:
            frmInbox.mnuTimeScale 10
            
        Case ID_CALENDAREVENT_5:
            frmInbox.mnuTimeScale 5
            
        Case Else
            AbiertoFormulario = True
            
            Abrir_Formularios Control_Id
            
            
    End Select
    
    
    If AbiertoFormulario Then
        AbiertoFormulario = False
        'mOTIVO... no lo se
        'Pero si lo vamos cambiando funciona
        If True Then
            If Me.DockingPaneManager.Panes(1).Enabled <> 3 Then
                Me.DockingPaneManager.Panes(1).Enabled = PaneEnabled
                Me.DockingPaneManager.Panes(2).Enabled = PaneEnabled
    
                'frmPaneCalendar.DatePicker.Enabled = True
                
                DockingPaneManager.RedrawPanes
                
                
            Else
                Me.DockingPaneManager.Panes(1).Enabled = PaneEnableActions
                Me.DockingPaneManager.Panes(2).Enabled = PaneEnableActions
                 
            End If
            DockingPaneManager.NormalizeSplitters
        End If
    End If
    
End Sub



Private Sub CommandBars_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
        Dim Control As CommandBarControl, ControlItem As CommandBarControl
        
        If TypeOf CommandBar Is RibbonBackstageView Then
            Debug.Print "RibbonBackstageView"
        End If
        
        Set Control = CommandBar.FindControl(, IDS_ARRANGE_BY)
        If Not Control Is Nothing Then
            Dim Index As Long
            Index = Control.Index
            Control.visible = False
            
            Do While Index + 1 <= CommandBar.Controls.Count
                Set ControlItem = CommandBar.Controls.Item(Index + 1)
                If ControlItem.Id = IDS_ARRANGE_BY Then
                    ControlItem.Delete
                Else
                    Exit Do
                End If
            Loop
            
'            Dim CurrentColumn As ReportColumn
'            For Each CurrentColumn In frmInbox. wndReportControl.Columns
'                Set ControlItem = CommandBar.Controls.Add(xtpControlButton, ID_REPORTCONTROL_COLUMN_ARRANGE_BY, CurrentColumn.Caption)
'                ControlItem.Parameter = CurrentColumn.ItemIndex
'                If Not frmInbox. wndReportControl.SortOrder.IndexOf(CurrentColumn) = -1 Then
'                    ControlItem.Checked = True
'                End If
'                If Not CurrentColumn.Visible Then
'                    ControlItem.Visible = False
'                End If
'            Next
        
        End If
End Sub

Private Sub CommandBars_SpecialColorChanged()
    Me.BackColor = CommandBars.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Private Sub CommandBars_ToolBarVisibleChanged(ByVal ToolBar As XtremeCommandBars.ICommandBar)
     'Debug.Print ToolBar.BarID
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
        
    On Error Resume Next
    
    
    
    Select Case Control.Id
        Case ID_VIEW_STATUSBAR:
            'Control.Checked = CommandBars.StatusBar.Visible
        
        
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_WEEK:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_MONTH:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarMonthView, True, False)
        
        Case ID_OPTIONS_ANIMATION:
            'Control.Checked = CommandBars.ActiveMenuBar.EnableAnimation
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
             '   Dim newFontHeight As Integer
             '   newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
             '   Control.Checked = IIf(RibbonBar.FontHeight = newFontHeight, True, False)
                
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
              '  Control.Checked = CommandBars.PaintManager.AutoResizeIcons

        Case ID_RIBBON_EXPAND:
            'Control.Visible = RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            'Control.Visible = Not RibbonBar.Minimized
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    If (Action = PaneActionSplitterResized) Then
        DockingPaneManager.RecalcLayout
        
        ' Save MRUShortcutBarWidth
        If (frmShortBar.ScaleWidth > MinimizedShortcutBarWidth And Container.Container.Type = PaneTypeSplitterContainer) Then
            'Debug.Print frmShortBar.ScaleWidth
            MRUShortcutBarWidth = frmShortBar.ScaleWidth
        End If
    Else
        If (Action = PaneActionSplitterResized) Then Debug.Print "Resizing "
    End If
End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Tag = PANE_SHORTCUTBAR Then
        Item.Handle = frmShortBar.hwnd
    ElseIf Item.Tag = PANE_REPORT_CONTROL Then
        Item.Handle = frmInbox.hwnd
    End If
End Sub

Private Sub Form_Activate()


    If PrimeraVez Then
        PrimeraVez = False
        
        
        'Para este usuario y esta empresa unos avlores al usuario
        If vEmpresa.codempre > 0 Then vUsu.FijarOtrosValoresUsuario
        
        Espera 0.25
        
        HACER_CommandBars_Execute id_acciones_inicio
        
        DoEvent2
        
        

    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaDatosMenusDemas(DesdeLoad As Boolean)
Dim AntiguoTab As Integer
    
    
    Screen.MousePointer = vbHourglass
    AntiguoTab = -1
    If RibbonSeHaCreado Then
        If Not RibbonBar.SelectedTab Is Nothing Then AntiguoTab = RibbonBar.SelectedTab.Id
    End If
    CreateRibbon
    Screen.MousePointer = vbHourglass
    CreateBackstage
    Screen.MousePointer = vbHourglass
    CreateRibbonOptions
    
    
    If Not DesdeLoad Then vEmpresa.LeerDatos
    
    
    Screen.MousePointer = vbHourglass
    CargaMenu AntiguoTab
    CreateStatusBar
    Screen.MousePointer = vbHourglass
    PonerCaption
    CreateCalendarTabOriginal
    RibbonSeHaCreado = True
End Sub






Public Sub CambiarEmpresa_(QueEmpresa As Integer)
Dim cur As Integer

    Screen.MousePointer = vbHourglass
    Me.Hide
    CambiarEmpresa2 QueEmpresa
    Me.Show
       DoEvents
       Screen.MousePointer = vbHourglass
    AccionesIncioAbrirProgramaEmpresa
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub CambiarEmpresa2(QueEmpresa As Integer)
Dim RB As RibbonBar
    CadenaDesdeOtroForm = vUsu.Login & "|" & vEmpresa.codempre & "|"
        
    
        
    Set vUsu = New Usuario
    vUsu.Leer RecuperaValor(CadenaDesdeOtroForm, 1)
    
    
    vUsu.CadenaConexion = "ariges" & IIf(QueEmpresa = 0, "", QueEmpresa)
    
    vUsu.LeerFiltros "ariges", 301 ' asientos
    vUsu.LeerFiltros "ariges", 401 ' facturas de cliente
    
    AbrirConexion  'Usu.CadenaConexion
    
    Set vEmpresa = New Cempresa
    Set vParam = New Cparametros
    
    vEmpresa.LeerDatos
    vParam.Leer
    
    UltimoEmpresaLogada False, vUsu.CadenaConexion
    
    PonerCaption
    
    Screen.MousePointer = vbHourglass
   CargaDatosMenusDemas True
   frmPaneContacts.SeleccionarNodoEmpresa vEmpresa.codempre
   pageBackstageHelp.Label9.Caption = vEmpresa.nomempre
   pageBackstageHelp.tabPage(0).visible = False
   pageBackstageHelp.tabPage(1).visible = False
   frmInbox.OpenProvider
   Set RB = RibbonBar
   RB.Minimized = False
   RB.RedrawBar
   
   
  
   'FALTA
   ' vControl.UltEmpre = vUsu.CadenaConexion
   '    vControl.Grabar
    
    
    
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()

    'Cargamos librerias de icinos de los forms
    If vUsu.Skin >= 0 Then frmIdentifica.pLabel "Carga DLL"
    CargaIconosDlls
   
     T1 = Timer  '
   
    CommandBarsGlobalSettings.App = App
   
    If vUsu.Skin >= 0 Then frmIdentifica.pLabel "Leyendo menus usuario"
    CargaDatosMenusDemas True
  
    ShowEventInPane = False
       
    FontSizes(0) = 0
    FontSizes(1) = 11
    FontSizes(2) = 13
    FontSizes(3) = 16
               
    DockingPaneManager.SetCommandBars Me.CommandBars
              
    Set frmShortBar = New frmShortcutBar2
    Set frmInbox = New frmInbox
        
    Dim A As Pane, B As Pane, C As Pane, D As Pane
    
    If vUsu.Skin >= 0 Then frmIdentifica.pLabel "Creando paneles"
    Set A = DockingPaneManager.CreatePane(PANE_SHORTCUTBAR, 170, 120, DockLeftOf, Nothing)
    A.Tag = PANE_SHORTCUTBAR
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    
    Set B = DockingPaneManager.CreatePane(PANE_REPORT_CONTROL, 700, 400, DockRightOf, A)
    B.Tag = PANE_REPORT_CONTROL
   
    DockingPaneManager.Options.HideClient = True
    PonerTabPorDefecto -1
    
    Set CommandBars.Icons = CommandBarsGlobalSettings.Icons
    LoadIcons
    
    DockingPaneManager.RecalcLayout
    MRUShortcutBarWidth = frmShortBar.ScaleWidth
   
   
   
   
    'En funcion
    ' ID_OPTIONS_STYLEBLUE2010  ID_OPTIONS_STYLESILVER2010    ID_OPTIONS_STYLEBLACK2010
    If vUsu.Skin >= 0 Then frmIdentifica.pLabel "Carga skin"
    Screen.MousePointer = vbHourglass
    
    If vUsu.Skin = 3 Then
        Cad = ID_OPTIONS_STYLEBLACK2010
    Else
        If vUsu.Skin = 2 Then
            Cad = ID_OPTIONS_STYLESILVER2010
        Else
            Cad = ID_OPTIONS_STYLEBLUE2010
        End If
    End If
    CommandBars.FindControl(, Cad, , True).Execute
    
    PrimeraVez = True

    
End Sub


Private Sub CargaIconosDlls()
Dim TamanyoImgComun As Integer

'    ImageList1.ImageHeight = 48
'    ImageList1.ImageWidth = 48
'    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 48
'
'
'    ImageList2.ImageHeight = 16
'    ImageList2.ImageWidth = 16
'    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 16
'
'    ImageListPPal48.ImageHeight = 48
'    ImageListPPal48.ImageWidth = 48
'    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 8, 48
'
'
'    ImageListPpal16.ImageHeight = 16
'    ImageListPpal16.ImageWidth = 16
'    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 9, 16
'
'
'
'    ImgListComun2.ListImages.Clear
'    imgListComun_BN2.ListImages.Clear
'    imgListComun_OM2.ListImages.Clear
'
'        TamanyoImgComun = 24
'
'        ImgListComun2.ImageHeight = TamanyoImgComun
'        ImgListComun2.ImageWidth = TamanyoImgComun
'        GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 2, TamanyoImgComun  'antes icolistcon
'
'
'
'        '++
'        imgListComun_BN2.ImageHeight = TamanyoImgComun
'        imgListComun_BN2.ImageWidth = TamanyoImgComun
'        GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 3, TamanyoImgComun
'
'        imgListComun_OM2.ImageHeight = TamanyoImgComun
'        imgListComun_OM2.ImageWidth = TamanyoImgComun
'        GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 4, TamanyoImgComun
'
'
'    imgListComun16.ImageHeight = 16
'    imgListComun16.ImageWidth = 16
'    GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 5, 16
'
'    GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 6, 16
'    GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 7, 16
'

End Sub

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal OP As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = OP
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)
   
    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub



Public Sub ExpandButtonClicked()
   
    
    
    Dim A As Pane
    Set A = DockingPaneManager.FindPane(PANE_SHORTCUTBAR)
    
    Dim ShortcutBarMinimized As Boolean
    ShortcutBarMinimized = frmShortBar.ScaleWidth <= MinimizedShortcutBarWidth
    
    Dim NewWidth As Long
    If (ShortcutBarMinimized) Then
        NewWidth = MRUShortcutBarWidth
    Else
        NewWidth = MinimizedShortcutBarWidth
        frmShortBar.wndShortcutBar.PopupWidth = MRUShortcutBarWidth
    End If
        
    
    ' Set Size of Pane
    A.MinTrackSize.Width = NewWidth
    A.MaxTrackSize.Width = NewWidth
        
    DockingPaneManager.RecalcLayout
    DockingPaneManager.NormalizeSplitters
    DockingPaneManager.RedrawPanes
    
    ' Restore Constraints
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    A.MaxTrackSize.Width = 32000
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not (pageBackstageInfo Is Nothing) Then Unload pageBackstageInfo
    If Not (pageBackstageHelp Is Nothing) Then Unload pageBackstageHelp
    'If Not (pageBackstageSend Is Nothing) Then Unload pageBackstageSend
    
    'close all sub forms
    On Error Resume Next
    Dim i As Long
    For i = Forms.Count - 1 To 1 Step -1
        
        Unload Forms(i)
    Next
    
    
    GuardarDatosUltimaTab
  
  
    'Cerrar Conexion
    Set conn = Nothing
    Set ConnConta = Nothing
    End
End Sub



Private Sub GuardarDatosUltimaTab()
Dim i As Integer
    i = RibbonBar.SelectedTab.Id
    If i = ID_TAB_CALENDAR_HOME Then Exit Sub 'no guardo este tab
    If i <> vUsu.TabPorDefecto Then
        vUsu.TabPorDefecto = i
        vUsu.GuardarTabPorDefecto
    End If
End Sub


Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, Id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, Id, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddButton = Control
    
End Function

Private Sub CommandBars_Resize()
    
    On Error Resume Next
    
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
   ' If Not PrimeraVez Then Exit Sub
    
    CommandBars.GetClientRect Left, top, Right, Bottom
    
End Sub

Private Sub LoadIcons()
Dim T() As Variant
Dim Contador As Integer
Dim CualBmp As Byte '1-2-3-4 con
Dim MaxIco As Byte
Dim fin As Boolean
Dim tamayoMaximoArrayPpal As Integer
Dim i As Integer
Dim K As Byte


    If False Then
        Exit Sub
    End If
    
    
    CommandBars.Icons.RemoveAll
    SuiteControlsGlobalSettings.Icons.RemoveAll
    ReportControlGlobalSettings.Icons.RemoveAll

    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\help.png", ID_APP_ABOUT, xtpImageNormal



    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookcalicons.png", _
            Array(ID_GROUP_NEW_APPOINTMENT, ID_GROUP_NEW_MEETING, ID_GROUP_NEW_ITEMS, ID_GROUP_GOTO_TODAY, _
            ID_GROUP_GOTO_NEXT7DAYS, ID_GROUP_ARRANGE_DAY, ID_GROUP_ARRANGE_WORK_WEEK, ID_GROUP_ARRANGE_WEEK, _
            ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_SCHEDULE_VIEW, ID_GROUP_MANAGE_CALENDARS_OPEN, ID_GROUP_MANAGE_CALENDARS_GROUPS, _
            1, 1, 1, 1), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\RibbonMinimize.png", _
            Array(ID_RIBBON_MINIMIZE, ID_RIBBON_EXPAND), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\Search.png", _
            ID_SEARCH_ICON, xtpImageNormal
            
    If False Then
                 '------------------------------------------------------------------------------------------------------------------------
                 '------------------------------------------------------------------------------------------------------------------------
                 '------------------------------------------------------------------------------------------------------------------------
                         
           '       CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonslarge.png", _
                         Array(ID_GROUP_MAIL_NEW_NEW, ID_GROUP_MAIL_NEW_NEW_ITEMS, ID_GROUP_MAIL_DELETE_DELETE, ID_GROUP_MAIL_RESPOND_REPLY, _
                         ID_GROUP_MAIL_RESPOND_REPLY_ALL, ID_GROUP_MAIL_RESPOND_FORWARD, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
                         
                         
              '    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", _
                         Array(ID_GROUP_MAIL_DELETE_CLEANUP, ID_GROUP_MAIL_DELETE_JUNK, ID_GROUP_MAIL_RESPOND_MEETING, ID_GROUP_MAIL_RESPOND_IM, _
                         ID_GROUP_MAIL_RESPOND_MORE, ID_GROUP_MAIL_TAGS_UNREAD, ID_GROUP_MAIL_TAGS_CATEGORIZE, ID_GROUP_MAIL_TAGS_FOLLOWUP, ID_GROUP_MAIL_FIND_ADDRESSBOOK, _
                         ID_GROUP_MAIL_FIND_FILTER, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
                 
                 '    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookpane.png", _
                         Array(ID_SWITCH_NORMAL, ID_SWITCH_CALENAR_AND_TASK, ID_SWITCH_CALENDAR, ID_SWITCH_CLASSIC, ID_SWITCH_READING), xtpImageNormal
                         
                     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
                         Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
                         SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
                     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_24x24.bmp", _
                         Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
                         SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
                         
                  '   CommandBars.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
                         Array(ID_QUICKSTEP_REPLAY_DELETE, ID_QUICKSTEP_TO_MANAGER, ID_QUICKSTEP_MOVE_TO, ID_QUICKSTEP_CREATE_NEW, ID_QUICKSTEP_TEAM_EMAIL, ID_QUICKSTEP_DONE), xtpImageNormal
                         
                     ReportControlGlobalSettings.Icons.LoadBitmap App.Path & "\styles\bmreport.bmp", _
                     Array(COLUMN_MAIL_ICON, COLUMN_IMPORTANCE_ICON, COLUMN_CHECK_ICON, RECORD_UNREAD_MAIL_ICON, RECORD_READ_MAIL_ICON, _
                         RECORD_REPLIED_ICON, RECORD_IMPORTANCE_HIGH_ICON, COLUMN_ATTACHMENT_ICON, COLUMN_ATTACHMENT_NORMAL_ICON, _
                         RECORD_IMPORTANCE_LOW_ICON), xtpImageNormal
                         
                         
                     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\suministro-inmediato-informacion.bmp", 1, xtpImageNormal
                 
                
        End If
  
      
        
    
    
    
    
     
    'Deberiamos cargar un array con unos(1) de longitud 143
    ' y en funcion del valor del campo imagen en el punto de menu correspondiente
    ' lo pondremos en el array.
    ' Ejemplo    303 Extractos  Campo imagen: 87
    ' quiere decir que en el campo 87 del array sustituieremos el 1 por el 303


'
    
    'Cad linea son 15
    T = Array(2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1)
    
    
    
    'Vamos a necesitar un array de arrays
    'Cada bloque de cuatro son:
    '   11|9|5|11|  Total: 40-> Necesito, de momento 10 bloques
    Dim J As Integer
    
    tamayoMaximoArrayPpal = 35
    ReDim MacroArray(tamayoMaximoArrayPpal)
    
    Dim Arry()
    
    For J = 0 To tamayoMaximoArrayPpal
        'cual:      1-bmreport2013_16   12
        '           2-mail_16x16        10
        '           3-quickstepsgallery 6
        '           4-reporticonssmall  12
        MaxIco = CByte(RecuperaValor("11|9|5|11|", (CInt(J) Mod 4) + 1))
        
        K = (J Mod 3)
        If K = 0 Then
            MaxIco = 11
        ElseIf K = 1 Then
            MaxIco = 5
        Else
            MaxIco = 9
        End If
            
        ReDim Arry(MaxIco)
        For K = 0 To MaxIco
            Arry(K) = CLng(1)
        Next
        MacroArray(J) = Arry()
       
    Next
    
    Dim Rn2 As ADODB.Recordset
    Cad = "select * from menus where aplicacion='ariges' and padre>0   order by imagen desc,padre,codigo "
    
    'Leeemos un unico RS
    Set Rn2 = New ADODB.Recordset
    
    Rn2.Open Cad, conn, adOpenForwardOnly, adLockOptimistic  'NO PUEDE SER EOF
    Contador = 0
    'Primero los grandes
    fin = False
    Do
        If Rn2.EOF Then
            fin = True
        Else
        
        
            If Rn2!imagen = 0 Then
                fin = True
            Else
                If Contador > 99 Then
                    fin = True
                Else
                    T(Contador) = Rn2!Codigo
                    Contador = Contador + 1
                End If
                Rn2.MoveNext
            End If
            
        End If
    Loop Until fin
        
    'Llenamos losa agrupados2
    ' Seran botones que contienen botonoes. Como mucho 35. SE pueden repitr, obivamente
    For Contador = 100 To 135
        T(Contador) = 150 + (Contador - 100)
    Next
    'Cargmos iconos 32x32
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", T, xtpImageNormal
    
           
               
           
    fin = Rn2.EOF 'No deberia
    
    If Not fin Then
        'Iconos pequeños
        '---------------------------------------------------------
        ' Tenemos 3 bms con iconos
        'Con lo cual iremos cargando a medida que necesitemos
        'cual:      1-bmreport2013_16   12
        '           2-mail_16x16        10
        '           3-quickstepsgallery 6
        '           4-reporticonssmall  12
        Contador = 1000

        
        
        
        
        
        
        J = -1
        Do
            
            If Contador > MaxIco Then
            
                
                J = J + 1
                'Nuevo Array
              
                '1-bmreport2013_16   12     2-mail_16x16  10     3-quickstepsgallery 6  4-reporticonssmall  12
                MaxIco = CByte(RecuperaValor("11|9|5|11|", CInt(CualBmp) + 1))
                
                K = (J Mod 3)
                If K = 0 Then
                    'Pares
                    MaxIco = 11
                ElseIf K = 1 Then
                    MaxIco = 5
                Else
                    MaxIco = 9
                End If
                Contador = 0
                
                
            End If
         
            If Rn2.EOF Then
                fin = True
                
                
                While Contador <= MaxIco
                    T(Contador) = 1
                    Contador = Contador + 1
                Wend
                
            Else
            
                
                MacroArray(J)(Contador) = CLng(Rn2!Codigo)
                T(Contador) = Rn2!Codigo
                Contador = Contador + 1
                Rn2.MoveNext
            End If
            
        Loop Until fin
    End If
    Rn2.Close
    Set Rn2 = Nothing
    
    
    'Cargamos
    For i = 0 To tamayoMaximoArrayPpal
        T = MacroArray(i)
        K = (i Mod 3)
        If K = 0 Then
            CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", T, xtpImageNormal
        ElseIf K = 1 Then
            CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", T, xtpImageNormal
        Else
            CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal
        End If
    Next i
    
    
    
        For i = 1 To 17
            SuiteControlsGlobalSettings.Icons.LoadIcon App.Path & "\styles\TreeView\icon" & i & ".ico", i, xtpImageNormal
        Next i
        
   'T = Array(211, 213, 233, 234, 235, 236, 237, 238, 301, 302)
        
   'CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal
    
'    Dim N()
'    N = Array(211, 213, 233, 234, 235, 236, 237, 238, 301, 302, 303, 304)
'
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", N, xtpImageNormal
'    N = MacroArray(5)
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", N, xtpImageNormal
'    N = MacroArray(4)
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", N, xtpImageNormal
End Sub


Private Sub LoadLeerPngSobreArray(QuePNG As Byte, ByRef TArray)
Dim AuxArray()
        'Iconos pequeños
        '---------------------------------------------------------
        ' Tenemos 3 bms con iconos
        'Con lo cual iremos cargando a medida que necesitemos
        'cual:      1-bmreport2013_16   12
        '           2-mail_16x16        10
        '           3-quickstepsgallery 6
        '           4-reporticonssmall  12
        

        AuxArray = TArray
        Select Case QuePNG
        Case 2
            CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", TArray, xtpImageNormal
        
        Case 3
            CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", TArray, xtpImageNormal
        
        Case 4

            CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", TArray, xtpImageNormal
        Case Else

            CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\bmreport2013_16.png", AuxArray, xtpImageNormal
        End Select




End Sub


Private Sub SaveRibbonBarToXML()
    Dim Px As PropExchange
    Set Px = XtremeCommandBars.CreatePropExchange()
    
    Px.CreateAsXML False, "Settings"
        
    Dim Options As StateOptions
    Set Options = CommandBars.CreateStateOptions()
    Options.SerializeControls = True
        
    CommandBars.DoPropExchange Px.GetSection("CommandBars"), Options
    
    Px.SaveToFile App.Path & "\Layout.xml"
    
End Sub


'
'Private Function CreateQuickStepGallery() As CommandBarGalleryItems
'
'    Dim GalleryItems As CommandBarGalleryItems
'    Set GalleryItems = CommandBars.CreateGalleryItems(ID_GALLERY_QUICKSTEP)
'
'    GalleryItems.ItemWidth = 120
'    GalleryItems.ItemHeight = 20
'
'    GalleryItems.AddItem ID_QUICKSTEP_MOVE_TO, "Move To: ?"
'    GalleryItems.AddItem ID_QUICKSTEP_TO_MANAGER, "To Manager"
'    GalleryItems.AddItem ID_QUICKSTEP_TEAM_EMAIL, "Team E-mail"
'    GalleryItems.AddItem ID_QUICKSTEP_DONE, "Done"
'    GalleryItems.AddItem ID_QUICKSTEP_REPLAY_DELETE, "Reply & Delete"
'    GalleryItems.AddItem ID_QUICKSTEP_CREATE_NEW, "Create New"
'
'    GalleryItems.Icons = CommandBarsGlobalSettings.Icons
'
'    Set CreateQuickStepGallery = GalleryItems
'
'End Function

Private Sub CommandBars_ControlNotify(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal Code As Long, ByVal NotifyData As Variant, Handled As Variant)
   
    If (Code = XTP_BS_TABCHANGED) Then

        
    End If
End Sub


Private Sub CreateBackstage()

    
    Dim RibbonBar As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
    Dim BackstageView As RibbonBackstageView
    Set BackstageView = CommandBars.CreateCommandBar("CXTPRibbonBackstageView")
    
    BackstageView.SetTheme xtpThemeRibbon


    CommandBars.Icons.LoadBitmap App.Path & "\styles\BackstageIcons.png", _
    Array(1, 1, 1002, 1, 1, ID_APP_EXIT), xtpImageNormal

    Set RibbonBar.AddSystemButton.CommandBar = BackstageView
    
    'BackstageView.AddCommand ID_FILE_SAVE, "Cambiar empresa"
    'BackstageView.AddCommand ID_FILE_SAVE_AS, "Personalizar"
    'BackstageView.AddCommand ID_FILE_OPEN, "Open"
    'BackstageView.AddCommand ID_FILE_CLOSE, "Close"
    
    'If (pageBackstageInfo Is Nothing) Then Set pageBackstageInfo = New pageBackstageInfo
    'If (pageBackstageSend Is Nothing) Then Set pageBackstageSend = New pageBackstageSend
    If (pageBackstageHelp Is Nothing) Then Set pageBackstageHelp = New pageBackstageHelp
    
    Dim ControlInfo As RibbonBackstageTab
    Set ControlInfo = BackstageView.AddTab(1000, "Info", pageBackstageHelp.hwnd)
    
    'BackstageView.AddTab 1002, "Empresas", pageBackstageSend.hwnd

    ' Los menus de informacion...
    'BackstageView.AddTab 1001, "Acerca de", pageBackstageInfo.hwnd
    
    
    
    
    
    
    
    
    
    
    'BackstageView.AddCommand ID_FILE_OPTIONS, "Options"
    BackstageView.AddCommand ID_APP_EXIT, "Salir"
    
    ControlInfo.DefaultItem = True
    

End Sub




Private Sub CreateCalendarTabOriginal()

    Dim TabCalendarHome As RibbonTab
    Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup

    
    Dim Control As CommandBarControl
    Dim ControlNew_NewItems As CommandBarPopup
    Dim ControlArrange_Month As CommandBarPopup
    Dim ControlManage_Open As CommandBarPopup
    Dim ControlManage_Groups As CommandBarPopup
    Dim ControlShare_Publish As CommandBarPopup
           
    Dim PopupBar As CommandBar
    
    Set TabCalendarHome = RibbonBar.InsertTab(14, "Agenda")
    TabCalendarHome.Id = ID_TAB_CALENDAR_HOME
 
    Set GroupNew = TabCalendarHome.Groups.AddGroup("&Nueva", ID_GROUP_NEW)
        
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Evento")
    Control.Enabled = False
    Control.visible = False
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_MEETING, "&Cita")
    Control.Enabled = True
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_SHARE_SHARE, "&Imprimir")
    Control.Enabled = True
    
    
    
    '------------------------------------
    'Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "Evento")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "E&vento todo el dia")
    '    Control.BeginGroup = True
    'ControlNew_NewItems.KeyboardTip = "V"
    
    Set GroupGoTo = TabCalendarHome.Groups.AddGroup("I&r a", ID_GROUP_GOTO)
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_TODAY, "&Hoy")
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_NEXT7DAYS, "Próximos &7 dias ")
    GroupGoTo.ShowOptionButton = True
    GroupGoTo.ControlGroupOption.Caption = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.ToolTipText = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.DescriptionText = "Ir a fecha especificada."
    
    Set GroupArrange = TabCalendarHome.Groups.AddGroup("Vista", ID_GROUP_ARRANGE2)
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_DAY, "&Dia vista")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WORK_WEEK, "Samana &trabajo")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WEEK, "Sema&na vista")
    Set ControlArrange_Month = GroupArrange.Add(xtpControlSplitButtonPopup, ID_GROUP_ARRANGE_MONTH, "Mes")
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_LOW, "Ver detalle")
            Control.ToolTipText = "Muestra solo eventos todo el dia."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_MEDIUM, "Detalle &Medio")
            Control.ToolTipText = "Eventos todo el dia y si esta libre el dia o tiene eventos."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_HIGH, "Detalle &Alto")
            Control.ToolTipText = "Muestra todo."
            Control.DescriptionText = Control.ToolTipText

'    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_SCHEDULE_VIEW, "Schedule View")
'    GroupArrange.ShowOptionButton = True
'    GroupArrange.ControlGroupOption.Caption = "Calendar Options"
'    GroupArrange.ControlGroupOption.ToolTipText = "Calendar Options"
'    GroupArrange.ControlGroupOption.DescriptionText = "Change the settings for calendars, meetings and time zones."
'
'
  
    
End Sub





Private Sub CreateRibbon()
    Dim RibbonBar As RibbonBar
    
    If RibbonSeHaCreado Then Exit Sub
        
    
    
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    
    RibbonBar.AllowQuickAccessCustomization = False
    RibbonBar.ShowQuickAccessBelowRibbon = False
    RibbonBar.ShowGripper = False
    
    RibbonBar.AllowMinimize = False
    RibbonBar.AddSystemButton
    
    RibbonBar.SystemButton.IconId = ID_SYSTEM_ICON
    RibbonBar.SystemButton.Caption = "&Menu"
    RibbonBar.SystemButton.Style = xtpButtonCaption
End Sub

Private Sub CreateRibbonOptions()

    CommandBars.EnableActions
    If RibbonSeHaCreado Then Exit Sub
    
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue", "Office 2010 Blue", "Office 2010 Blue", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver", "Office 2010 Silver", "Office 2010 Silver", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black", "Office 2010 Black", "Office 2010 Black", "Themes"

    Dim Control As CommandBarControl, ControlAbout As CommandBarControl
    Dim ControlPopup As CommandBarPopup, ControlOptions As CommandBarPopup
         
    Set ControlOptions = RibbonBar.Controls.Add(xtpControlPopup, 0, "Opciones")
    ControlOptions.Flags = xtpFlagRightAlign
    
    Set Control = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Styles")
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black"
    
    Set ControlPopup = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Tamaño fuente", -1, False)
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_SYSTEM, "Sistema", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlRadioButton, ID_OPTIONS_FONT_NORMAL, "Normal", -1, False)
    Control.BeginGroup = True
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_LARGE, "Grande", -1, False
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_EXTRALARGE, "Extra grande", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_OPTIONS_FONT_AUTORESIZEICONS, "Ajustar Icons", -1, False)
    Control.BeginGroup = True
    
    'ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_RTL, "Right To Left"
    ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_ANIMATION, "Animation   "
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_MINIMIZE, "Minimizar la barra", False, "Muestra solo los titulos del menu principal.")
    Control.Flags = xtpFlagRightAlign
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_EXPAND, "Expandir la barra", False, "Muestra todos los elementos del menu.")
    Control.Flags = xtpFlagRightAlign
        
    Set ControlAbout = RibbonBar.Controls.Add(xtpControlButton, ID_APP_ABOUT, "&Acerca de")
    ControlAbout.Flags = xtpFlagRightAlign Or xtpFlagManualUpdate
    

        
End Sub








'*************************************************************************
'*************************************************************************
'*************************************************************************
'
'       CARGA menus en Ribbon
'
'




Public Sub CargaMenu(AntiguoTab As Integer)
Dim RN As ADODB.Recordset




    Set RN = New ADODB.Recordset
    Set Rn2 = New ADODB.Recordset
    On Error GoTo eCargaMenu
    

    If RibbonSeHaCreado Then RibbonBar.RemoveAllTabs
    
    Set RsMenusUsuarios = New ADODB.Recordset
    Cad = "select codigo,ver from menus_usuarios where aplicacion = " & DBSet("ariges", "T")
    Cad = Cad & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Cad = Cad & " ORDER by codigo"
    RsMenusUsuarios.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    
    
    
    Cad = "Select * from menus where aplicacion = 'ariges' and padre =0 and ocultarPorInstalacion=0 ORDER BY padre,orden "
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
    
        If Not BloqueaPuntoMenu(RN!Codigo, "ariges") Then
            Habilitado = True
             
            If Not MenuVisibleUsuarioPRE(DBLet(RN!Codigo), "ariges") Then Habilitado = False
                
            
                
            If Habilitado Then
                
                Select Case RN!Codigo
                Case 1
                    '1   "CONFIGURACION"
                    CargaMenuConfiguracion RN!Codigo
                Case Else
                    
                    CargaMenuGnerico RN!Codigo, RN!Descripcion
                    
                End Select
                
            End If
                                                 
        End If  'de habilitado el padre
    
        RN.MoveNext
    Wend
    RN.Close
                        
    PonerTabPorDefecto AntiguoTab
    
    
    
    
    'Hay algunos puntos de menu que tienen Captions distintos de pendidendo la
    'instalacion o configuracion
   
    
    '-- Descriptores especiales (Vrs 4.0.9)
    If vParamAplic.Descriptores Then
        'Creo que esto fue Morales, antes de duplicar proyecto
        'mnAlmTipoUnidad.Caption = "Formatos"
        CambiarCaption id_TiposUnidad, "Formatos"
                
        'mnTiposArticulos.Caption = "Modelos"
        CambiarCaption id_TiposArtículos, "Formatos"
        
        'mnAlmFamiliaArticulo.Caption = "Categorias Art."
        CambiarCaption id_TiposUnidad, "Categorias Art."
        
        
    End If
    If vParamAplic.Renting Then
        'mnManPrevFac2(4).Caption = "Facturación " & RentingLB & " y servicios"
        CambiarCaption id_FacturaciónRenting, "Facturación " & RentingLB & " y servicios"
        'mnManPrevFac2(3).Caption = "Previsión " & LCase(mnManPrevFac2(4).Caption)
        CambiarCaption id_PrevisiónRenting, "Previsión " & RentingLB & " y servicios"
    End If

    
    If vParamAplic.CartaPortes Then CambiarCaption id_FormasdeEnvío, "Transportistas"
    
    
    
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        'mnServicios(3).Caption = "Albarán de garantías"   para taxco
        CambiarCaption id_AlbaranesInternos, "Albarán de garantías"
        'mnMtoEuler(1).Caption = "Orden de taller"
        CambiarCaption id_AlbaránOrdendetrabajo, "Orden de taller"
        'mnMtoEuler(2).Caption = "Gestoria"
        CambiarCaption id_AlbaránServExterior, "Gestoria"
    End If
    
    
    
    
    
    
    
    
    
eCargaMenu:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    
    Set TabNuevo = Nothing
    Set GroupNew = Nothing
    Set Control = Nothing
    Set RN = Nothing
    Set Rn2 = Nothing
    Set RsMenusUsuarios = Nothing
End Sub


Private Sub CambiarCaption(id_Del_Menu_Codjecok As Long, NuevoCaption As String)
    If Not RibbonBar.CommandBars.ActiveMenuBar.Controls.Find(, id_Del_Menu_Codjecok) Is Nothing Then
        RibbonBar.CommandBars.ActiveMenuBar.Controls.Find(, id_Del_Menu_Codjecok).Caption = NuevoCaption
        RibbonBar.CommandBars.ActiveMenuBar.Controls.Find(, id_Del_Menu_Codjecok).ToolTipText = NuevoCaption
    End If
End Sub


Private Sub PonerTabPorDefecto(AntiguoTabSeleccionado As Integer)
Dim i As Integer
Dim J As Integer
Dim Anterior As Integer

    On Error Resume Next
    
    If AntiguoTabSeleccionado < 0 Then
        Anterior = vUsu.TabPorDefecto
    Else
        Anterior = AntiguoTabSeleccionado
    End If
    
    Cad = ""
    For i = 0 To RibbonBar.TabCount - 1
        J = RibbonBar.Tab(i).Id
        'Debug.Print J & " " & RibbonBar.Tab(i).Caption
        If J = Anterior Then
            
            RibbonBar.Tab(i).visible = True
            RibbonBar.Tab(i).Selected = True
            Set RibbonBar.SelectedTab = RibbonBar.Tab(i)
            Cad = "OK"
            Exit For
        End If
    Next
    If Cad = "" Then
        
        For J = RibbonBar.TabCount To 1 Step -1
            RibbonBar.Tab(J - 1).visible = True
            RibbonBar.Tab(J - 1).Selected = True
        Next J
    End If

    Err.Clear
End Sub

Private Sub CargaMenuConfiguracion(IdMenu As Integer)

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Configuracion")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
       
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariges' and padre =" & IdMenu & " and ocultarPorInstalacion=0 ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariges") Then
                Habilitado = True
    
    
                'Dependiendo de la empresa, podra no ver ciertas csas
                
    
    
                If Not MenuVisibleUsuarioPRE(DBLet(Rn2!Codigo), "ariges") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuarioPRE(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
                End If
           
                'FALTA###
                'If Rn2!Codigo = ID_ConfigurarBalances Then
                If False Then
'                    Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, Rn2!Codigo, Rn2!Descripcion)
'                    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConfigurarBalances1, "Balances")
'                    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConfigurarBalances2, "Ratios")
'                    If vUsu.Login = "root" Then Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConfigurarBalances3, "Personalizables")
'
                    'Personalizan
                    ControlNew_NewItems.Enabled = Habilitado
                Else
                    Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Control.Enabled = Habilitado
                End If
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
        
        'color Categorias  eventos
        If Not GroupNew Is Nothing Then
            Set Control = GroupNew.Add(xtpControlButton, 199, "Categorias calendario")
        End If
        Set GroupNew = Nothing
End Sub





Private Sub CargaMenuGnerico(IdMenu As Integer, Caption As String)
Dim GrupoAnt As String
Dim H As Integer
Dim IdTabMenu As Long
Dim CrearAgrupacion As Boolean

Dim PadreHabilitado As Byte  '101 Sin establecer    0. NO   1 Si

        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariges' and padre =" & IdMenu & " and ocultarPorInstalacion=0  ORDER BY grupo,agrupacion2,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        PadreHabilitado = 101
        If Not Rn2.EOF Then
        
            'Creamos la TAB
            Set TabNuevo = RibbonBar.InsertTab(IdMenu, Caption)
            TabNuevo.Id = CLng(IdMenu)
            IdTabMenu = CLng(IdMenu) * 10000
            
            Set ControlNew_NewItems = Nothing
            
        
           
            PadreHabilitado = IIf(MenuVisibleUsuarioPRE(DBLet(Rn2!Padre), "ariges"), 1, 0)
           
            
            GrupoAnt = "@#@#"
            While Not Rn2.EOF
            
               ' If Rn2!Codigo = 210 Then Stop
            
            
                If GrupoAnt <> DBLet(Rn2!Grupo, "T") Then
                    GrupoAnt = DBLet(Rn2!Grupo, "T")
                    If GrupoAnt = "" Then
                        Cad = ""
                    Else
                        Cad = Mid(GrupoAnt, 5)
                    End If
                    Set GroupNew = TabNuevo.Groups.AddGroup(Cad, IdTabMenu)
                End If
               If Not BloqueaPuntoMenu(Rn2!Codigo, "ariges") Then
                    If PadreHabilitado = 0 Then
                        Habilitado = False
                    Else
                        Habilitado = True
                        If Not MenuVisibleUsuarioPRE(DBLet(Rn2!Codigo), "ariges") Then Habilitado = False
                        
                    End If
                
                    
                    If DBLet(Rn2!agrupacion2, "T") <> "" Then
                        Cad = Mid(Rn2!agrupacion2, 5)
                        If Not ControlNew_NewItems Is Nothing Then
                            'Vemos si hay que crear otro
                            If ControlNew_NewItems.Caption = Cad Then
                                'Es el mismo
                                CrearAgrupacion = False
                            Else
                                CrearAgrupacion = True
                            End If
                         Else
                            CrearAgrupacion = True
                         End If
                         If CrearAgrupacion Then
                            idMenuIconoAgrupados2 = CInt(Val(Mid(Rn2!agrupacion2, 1, 3)))
                            If idMenuIconoAgrupados2 > 35 Then idMenuIconoAgrupados2 = 35
                            idMenuIconoAgrupados2 = 150 + idMenuIconoAgrupados2
                            Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, idMenuIconoAgrupados2, Cad)
                         End If
                        Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, CLng(Rn2!Codigo), Rn2!Descripcion)
                        
                        'Personalizan
                        ControlNew_NewItems.Enabled = Habilitado
                        CrearAgrupacion = True  'para que haga el nothing despues
                    Else
                        If CrearAgrupacion = True Then Set ControlNew_NewItems = Nothing
                        Set Control = GroupNew.Add(xtpControlButton, CLng(Rn2!Codigo), Rn2!Descripcion)
                        Control.Enabled = Habilitado
                    End If
                End If
                Rn2.MoveNext
            Wend
            
        End If
        Rn2.Close
        
        Set GroupNew = Nothing
End Sub








Private Sub CargaMenuDatosGenerales(IdMenu As Integer)

Dim B As Boolean


        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Datos generales")
        TabNuevo.Id = CLng(IdMenu)
        
End Sub

'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**********************************************************f****************************************************
Private Sub Abrir_Formularios(Accion As Long)
  
  

   
    Screen.MousePointer = vbHourglass
    Select Case Accion
    Case id_Empresa To id_Usuarios
        '
        ' --------------------------------
        '   Datos generales
        '
        AbrirFormDatosGenerales Accion
  
    Case id_acciones_inicio
        'Abre formularios de contailizacion , entrada albaranes...
        AccionesIncioAbrirProgramaEmpresa
  
    Case id_Marcas To id_Telematel
        '
        ' --------------------------------
        '   Almacenes
        '
        AbrirFormDatosGeneralesAlmacen Accion
  
    Case id_Actividades To id_TarifasTaxímetros
        '
        ' --------------------------------
        '   Datos basicos venta
        '
        AbrirFormDatosBasicosVenta Accion
    Case id_Ofertas To id_Alertas
        '
        ' --------------------------------
        '   Ofertas-Pedidos Ventas(4)
        '
        AbrirFormOfertasPedidos Accion
            
    Case id_Albaranes To id_InfCostes
        ' --------------------------------
        '   ALBARANES venta
        '
        AbrirFormAlbaranesVenta Accion

            
    Case id_Proveedores To id_InfProveedor_Marca_Familia
        ' --------------------------------
        '   Compras
        '
        AbrirFormsCompras Accion
        
   Case id_Trabajadores To id_Conceptos
        ' --------------------------------
        '   Compras        1001 1026
        '
        AbrirFormsAministracion Accion
     
   Case id_TiposdeContrato To id_FacturaciónRenting
        ' --------------------------------
        '   Mantenimientos 1100
        '
        AbrirFormsMantenimiento Accion
     
         
   Case id_NúmerosdeSerie To id_Facturación
        ' --------------------------------
        '   Reparaciones
        '
        AbrirFormsReparaciones Accion
     
   Case id_TiposAcciones To id_LlamadasClientes
        ' --------------------------------
        '   CRM
        '
        AbrirFormsCRM Accion
  
   Case id_Ordenesproducción To id_DeclaraciónalcoholAEAT
        ' --------------------------------
        '   PRoduccion 1400..
        '
        AbrirFormsProduccion Accion

   Case id_PantalladeVenta To id_ParámetrosTerminales
        ' --------------------------------
        '   TPV 150
        '
        AbrirFormsTPV Accion

 

   Case id_Accionesrealizadas To id_DeclaracionROPO
        ' --------------------------------
        '   Utilidades 2000
        '
        AbrirFormsUtilidades Accion




   End Select
     
   If Timer - UltimaLecturaReminders > 300 Then
        frmReminders.OnReminders xtpCalendarRemindersFire, Nothing
        If frmReminders.CuantosAvisos > 0 Then frmReminders.Show vbModal, Me
        CerrarAvisos
        UltimaLecturaReminders = Timer
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub CerrarAvisos()
    On Error Resume Next
    Unload frmReminders
    Err.Clear
End Sub









Private Sub AbrirMensajeBoxCodejock(QueMsg As Byte, OtrosDatos As String)

    
    Select Case QueMsg
    Case 0 To 10
        'Mensajes standard de la aplicacion
        
        
        
    Case 11
        
   '     Msg = "Importe descuadre: " & OtrosDatos
        'MuestraMsgCodejock2 "Ariconta6", "Existen asientos descuadrados", Msg, "Revise asientos", "", 0, False
         
   '     MuestraMsgAriadna "Ariconta6", "Existen asientos descuadrados", Msg, "Revise asientos", "", 0, False
    Case 12
        
   '     Msg = "Limite: " & UltimaFechaCorrectaSII(vParam.SIIDiasAviso, Now)
        
        'MuestraMsgCodejock2 "Ariadna software", "A.E.A.T.", Msg, "", "Ver facturas|Continuar|", 0, False
        MuestraMsgAriadna "Ariadna software", "A.E.A.T.", "Tiene facturas pendientes de comunicar al SII." & vbCrLf, "", "Continuar|Ver facturas|", 64, False
        
   '     Msg = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam ut elit sit amet quam tristique pretium ultricies ut nisl. Donec ultricies ante sodales bibendum dapibus. Suspendisse tincidunt tellus vel ante blandit, at lacinia enim tincidunt. Vivamus lectus libero, gravida eget augue a, lacinia finibus mi. Ut varius orci vehicula ipsum placerat cursus a quis nisi. Vivamus ut placerat lacus. Vivamus at euismod turpis, tincidunt commodo velit. Quisque a elementum erat. Nunc a malesuada urna, nec pretium ipsum. Nulla egestas metus vel lacus lobortis ullamcorper. Integer mollis tortor at velit pharetra aliquet sit amet et augue. Donec gravida imperdiet dui, a pretium leo pretium nec. In facilisis nunc arcu, non volutpat nibh ultricies in. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. In et accumsan ligula."

   ' Msg = Msg & " Proin gravida posuere convallis. Nunc eu diam in massa efficitur tristique vel porttitor metus. Nunc interdum urna metus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Curabitur in tempus ex. Etiam sagittis placerat neque, non iaculis lorem faucibus et. Sed posuere purus in malesuada condimentum. Morbi nec commodo odio. Vivamus a vehicula ante, eget pulvinar quam. Praesent tristique purus mi, quis feugiat velit lobortis id. Suspendisse potenti. Donec volutpat cursus imperdiet. Curabitur ornare porta sem. Nunc fringilla dolor orci, nec rhoncus sapien tincidunt vitae."

   ' Msg = Msg & " Proin auctor quis massa non ornare. Sed ut aliquam nulla. Donec ornare consequat neque in dapibus. Aliquam volutpat aliquet lectus vel scelerisque. Donec fermentum iaculis tempor. Aliquam erat volutpat. Nulla ut magna urna. Nam semper non leo sit amet eleifend. Praesent sed dictum quam. Aenean ullamcorper elit neque. Vestibulum vehicula nulla sit amet scelerisque varius. Nam blandit turpis sed dolor finibus vehicula et vitae leo. Vivamus egestas elit a iaculis facilisis. Sed mollis at velit ac finibus. In fermentum ipsum ac massa eleifend, ut suscipit augue tincidunt."

   ' Msg = Msg & " Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Pellentesque vel auctor sapien, ut bibendum massa. Phasellus tincidunt metus risus, eget accumsan arcu viverra ut. Nunc rhoncus augue at laoreet rutrum. Phasellus et tempus odio. In eleifend placerat justo, et posuere lorem. Interdum et malesuada fames ac ante ipsum primis in faucibus. Nullam vel erat in odio placerat tincidunt. Etiam finibus purus turpis, non volutpat tellus finibus consequat. Phasellus ultrices magna congue metus dignissim hendrerit. Suspendisse nec massa nisl. Quisque sit amet nunc quis ligula tempus vulputate vel eget nunc. Nunc tincidunt arcu est, nec pellentesque velit dapibus at. Donec metus ex, tempus non lorem in, facilisis congue metus. Etiam porttitor rutrum tortor. Maecenas sollicitudin lacinia ornare."
      
       ' MuestraMsgAriadna "Ariadna software", "", Msg, "", "Ver facturas|Continuar|", 0, False
        
        
       ' MsgBoxA Msg, vbQuestion
    End Select


    
End Sub



Private Sub AccionesIncioAbrirProgramaEmpresa()
Dim C As String
Dim Tiene_A_cancelar As Byte    '0: NO    1:  Cobros      2 : Pagos     3 Los dos
            
    vUsu.InicializaFiltrosEmpresa
    
    If vUsu.Nivel = 0 And vUsu.Id >= 0 Then
        'EmpresasQueYaHaComunicadoAsientosDescuadrados :  Para que solo lo haga una vez
                    
                   ' If HayQueMostrarEliminarRiesgoTalPag Then
'                        Screen.MousePointer = vbHourglass
'                        frmMensajes.Banco = IIf(Tiene_A_cancelar < 10, "", "N")
'                        frmMensajes.Tipo = IIf(Tiene_A_cancelar < 10, CStr(Tiene_A_cancelar), "1")
'                        frmMensajes.Opcion = 63
'                        frmMensajes.Show vbModal
'                   ' End If
                
        
    End If
    
    'DAVID enero2021
     CarpetaExportar
   
    
    'NUEVO 2017
    'Contabilizacion
    If False Then
    ComprobarFechaContabilizadas
    End If
    
    If vParamAplic.NumeroInstalacion = vbEuler Then

        CadenaDesdeOtroForm = "not fechaent is null AND 1"
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "scaalb", CadenaDesdeOtroForm, "1")
        'CadenaDesdeOtroForm = "1"
        If Val(CadenaDesdeOtroForm) > 0 Then
                        
     
            frmAvisosAlb.Show vbModal

        End If
        CadenaDesdeOtroForm = ""
    End If

    
 
    
    DoEvent2
    
    Screen.MousePointer = vbDefault
    
        
End Sub









'Establecer y fijar Skin
Public Sub EstablecerSkin(QueSkin As Integer)

    FijaSkin QueSkin

  ' Cargando el archivo del Skin
  ' ============================
    'frmPpal.SkinFramework1.LoadSkin Skn$, ""
    Me.SkinFramework1.ApplyWindow frmppalN.hwnd
    Me.SkinFramework1.ApplyOptions = Me.SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
    


    
End Sub



Private Sub CarpetaExportar()
On Error Resume Next
    If Dir(App.Path & "\Exportar", vbDirectory) = "" Then MkDir (App.Path & "\Exportar")
    If Err.Number <> 0 Then
        MsgBox "Error carpeta EXPORTAR." & vbCrLf & Err.Description, vbExclamation
        Err.Clear
    End If

End Sub


Private Function FijaSkin(numero)
    Me.SkinFramework1.ExcludeModule "crviewer9.dll"

  Select Case (numero)
 
           
            Case 1:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlue.ini"
            Case 2:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalSilver.ini"
            Case 3:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlack.ini"
                
                  
                
        
        
  End Select
    
End Function



Private Sub PonerCaption()
        Caption = "ARIGES 6    V-" & App.Major & "." & App.Minor & "." & App.Revision & "    usuario: " & vUsu.Nombre
        'Label33.Caption = "   " & vEmpresa.nomempre
End Sub


Public Sub OpcionesMenuInformacion(Id As Long)
    
   ' Select Case Id
   ' Case ID_Licencia_Usuario_Final_txt
   '     LanzaVisorMimeDocumento Me.hwnd, "c:\programas\Ariadna.rtf"
   ' Case ID_Licencia_Usuario_Final_web
   '     LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "Licenciadeuso.html"
   ' Case ID_Ver_Version_operativa_web
   '     LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "Ariconta-6.html"  ' "http://www.ariadnasw.com/clientes/"
   ' Case ID_Ver_CambiosVersion
   '     LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "Versiones.html"
   ' End Select
    
End Sub


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







'************************************************************************************
'************************************************************************************
'************************************************************************************
'
'                       Abir formularios
'
'************************************************************************************
'************************************************************************************
'************************************************************************************
'Abrir formularios Datos Generarles
Private Sub AbrirFormDatosGenerales(ByRef Accion As Long)
    
''case id_Empresa = 101                            'Configuración(1)
''case id_ParámetrosAplicación = 102              'Configuración(1)
''case id_TiposMovimiento = 104                   'Configuración(1)
''case id_TiposDocumentos = 105                   'Configuración(1)
''case id_Usuarios = 106                           'Configuración(1)


    Select Case Accion
    Case id_Empresa
        frmConfParamGral.Show vbModal
                    
    Case id_ParámetrosAplicación
     
        Load frmConfParamAplic
        frmConfParamAplic.Show vbModal
    
    Case id_TiposMovimiento
        
        frmConfTipoMov.Show vbModal   'tipomov+++
    
    Case id_TiposDocumentos
          frmConfParamRpt.Show vbModal
    Case id_Usuarios
        
        If vUsu.Nivel > 0 Then Exit Sub
        frmMantenusu2.Show vbModal
    End Select
End Sub



Private Sub AbrirFormDatosGeneralesAlmacen(ByRef Accion As Long)
  
  
    Select Case Accion
        Case id_Marcas          '= 201                              'Almacen(2)
            frmAlmMarcas.Show vbModal
        Case id_AlmacenesPropios '= 202                   'Almacen(2)
            frmAlmAlPropios.Show vbModal
        Case id_TiposUnidad         ' = 203                        'Almacen(2)
             frmAlmTipoUnidad.Show vbModal
        Case id_TiposArtículos  '= 204                     'Almacen(2)
            frmAlmTipoArticulo.Show vbModal
        Case id_Ubicaciones '= 205                         'Almacen(2)
            frmAlmUbicaciones.Show vbModal
        Case id_Familias '= 206                            'Almacen(2)
            frmAlmFamiliaArticulo.Show vbModal
        Case id_Categorias '= 207                          'Almacen(2)
            frmAlmCategorias.Show vbModal
        Case id_Artículos '= 208                           'Almacen(2)

            
                frmAlmArticulosGr.DatosADevolverBusqueda = ""
                frmAlmArticulosGr.Show vbModal
            
        
        Case id_Númerosdelote '= 209                     'Almacen(2)
            frmAlmNumLote.Show vbModal
        
        Case id_Telematel '= 238                           'Almacen(2)
            frmTelematMto.Show vbModal
        
        Case id_TraspasoAlmacenes, id_HcoTraspasoAlmacenes          '= 210  211                'Almacen(2)
            frmAlmTraspaso.EsHistorico = Accion = id_HcoTraspasoAlmacenes
            frmAlmTraspaso.hcoCodMovim = -1
            frmAlmTraspaso.Show vbModal
                   
        
        Case id_MovimientosAlmacén, id_HcoMovimientosAlmacén                '= 212,213                 'Almacen(2)
            frmAlmMovimientos.EsHistorico = Accion = id_MovimientosAlmacén
            frmAlmMovimientos.hcoCodMovim = -1 'No carga el form al abrir
            frmAlmMovimientos.Show vbModal
        
        
        Case id_MovimientosArtículos           '= 214               'Almacen(2)
            frmAlmMovimArticulos.Show vbModal
        Case id_MovimientosStockdesdeInv           '= 215        'Almacen(2)
            frmAlmMovArtSaldo.Show vbModal
        Case id_InfControlStock           '= 216                   'Almacen(2)
            AbrirListado3 27
        Case id_InfArtículosinactivos           '= 217             'Almacen(2)
            AbrirListado (15)
        Case id_InfArtículoscomponentes           '= 218           'Almacen(2)
            AbrirListado (11)
        Case id_InfValoraciónstocks          ' = 219               'Almacen(2)
            AbrirListado (17)
        Case id_InfStocksmax_min           '= 220                  'Almacen(2)
            AbrirListado (18)
        Case id_InfStocksaFecha           '= 221                  'Almacen(2)
            AbrirListado (19)
        Case id_InfStocksxmeses           '= 222                  'Almacen(2)
             AbrirListado3 4
        Case id_InfAlertasPPedido          ' = 223                'Almacen(2)
             AbrirListado3 26
        Case id_InfReposiciónAlmacén           '= 224              'Almacen(2)
            AbrirListado3 35
        Case id_InfStockmínimo           '= 225                    'Almacen(2)
            AbrirListado (100)
        Case id_MovimientosLotes           '= 226                   'Almacen(2)
            AbrirListado2 54
        Case id_TomadeInventario           '= 227                  'Almacen(2)
             AbrirListado 12
        Case id_EntradaExistencia           '= 228                  'Almacen(2)
            frmAlmInventarioGR.Show vbModal
        Case id_ListadoDiferencias           '= 229                 'Almacen(2)
            AbrirListado (13)
        Case id_Actualizardiferencias           '= 230              'Almacen(2)
            AbrirListado (14)
        Case id_ValoraciónStocksInv           '= 231              'Almacen(2)
            AbrirListado (16)
        Case id_RectificarúltimoInv, id_InventariarArtículo          '= 232              'Almacen(2)
                If vUsu.Nivel > 2 Then
                    MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
                Else
                    AbrirListado3 IIf(Accion = id_RectificarúltimoInv, 61, 3)
                End If
                   
        
        Case id_RecálculoPrStandard, id_RecálculoPrMedioP, id_RecálculoUltPrCompra                    '= 234  235  236             'Almacen(2)
        
            If vUsu.Nivel > 1 Then
                MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
                Exit Sub
            Else
                If Accion = id_RecálculoPrStandard Then
                    frmListado3.Opcion = 6
                ElseIf Accion = id_RecálculoPrMedioP Then
                    frmListado3.Opcion = 20
                Else
                    frmListado3.Opcion = 21
                End If
            End If
            frmListado3.Show vbModal
        
        Case id_HistóricoInventario           ' = 237                'Almacen(2)
            frmAlmHcoInvenGR.Show vbModal
        End Select


End Sub



Private Sub AbrirFormDatosBasicosVenta(ByRef Accion As Long)
 

  



    Select Case Accion
    Case id_Actividades '= 301                        'Datos Básicos Ventas(3)
        frmFacActividades.Show vbModal
    Case id_Zonas '= 302                              'Datos Básicos Ventas(3)
        frmFacZonas.Show vbModal
    Case id_Rutas '= 303                              'Datos Básicos Ventas(3)
        frmFacRutas.Show vbModal

    Case id_Portes ' = 304                             'Datos Básicos Ventas(3)
        frmFacPortes.Show vbModal
    Case id_Descuentosporcantidad '= 305            'Datos Básicos Ventas(3)
        frmFacDtoUd.Show vbModal
    Case id_FormasdeEnvío '= 306                    'Datos Básicos Ventas(3)
        frmFacFormasEnvio.Show vbModal
    Case id_FormasdePago '= 307                     'Datos Básicos Ventas(3)
        frmFacFormasPago.Show vbModal
    Case id_Bancospropios '= 308                     'Datos Básicos Ventas(3)
        frmFacBancosPropios.Show vbModal
    Case id_SituacionesEspeciales '= 309             'Datos Básicos Ventas(3)
        frmFacSituaciones.Show vbModal
    Case id_Agentes '= 310                            'Datos Básicos Ventas(3)
        frmFacAgentesCom.Show vbModal
    Case id_Clientesvarios '= 311                    'Datos Básicos Ventas(3)
        frmFacClientesV.Show vbModal
    Case id_Clientes '= 312                           'Datos Básicos Ventas(3)
        frmFacClientesGr.Show vbModal
    Case id_ClientesPotenciales '= 313               'Datos Básicos Ventas(3)
        frmFacClienPot.Show vbModal
    Case id_TiposdeCartas ' = 314                    'Datos Básicos Ventas(3)
        frmFacCartasOferta.Show vbModal
    Case id_Incidencias '= 315                        'Datos Básicos Ventas(3)
        frmIncidencias.Show vbModal
    Case id_ClientesInactivos '= 316                 'Datos Básicos Ventas(3)
         'Informe de Clientes Inactivos
        AbrirListadoOfer (46) '46: Informes Clientes Inactivos
    Case id_AltasClientes '= 317                     'Datos Básicos Ventas(3)
          'Informe de Altas de Nuevos Clientes
        AbrirListadoOfer (48) '48: Informes Altas Clientes
    Case id_EtiquetasdeClientes '= 318              'Datos Básicos Ventas(3)
          'Etiquetas de clientes
        AbrirListadoOfer (90) '90: Informe Etiquetas de Clientes
    Case id_CartasaClientes '= 319                  'Datos Básicos Ventas(3)
          AbrirListadoOfer (91) '91: Informe Cartas a Clientes
    Case id_Etiquetasdebultos '= 320                'Datos Básicos Ventas(3)
        AbrirListado 95
    Case id_InfTeléfonosxCliente '= 321            'Datos Básicos Ventas(3)
          AbrirListado3 41
    Case id_InfCuotastelefonía '= 322               'Datos Básicos Ventas(3)
         AbrirListado3 48
    Case id_TarifasVenta '= 331                      'Datos Básicos Ventas(3)
        'Tarifas Venta
        frmFacTarifas.Show vbModal
    Case id_ListaPrecios '= 332                      'Datos Básicos Ventas(3)
        'Listado Precios
        frmFacTarifasPrecios.Show vbModal
    Case id_PreciosEspeciales '= 333                 'Datos Básicos Ventas(3)
       'Precios especiales
        frmFacPreciosEspecial.CadenaSituarData = ""
        frmFacPreciosEspecial.Show vbModal
    
    Case id_Promociones '= 334                        'Datos Básicos Ventas(3)
         'PROMOCIONES
        frmFacPromociones.Show vbModal
    Case id_DtosFamilia_Marca '= 335                'Datos Básicos Ventas(3)
         'Dots familia marca
        frmFacDtosFamMarca.Show vbModal
    Case id_DtosxActividad '= 336                  'Datos Básicos Ventas(3)
         'dtos por activiad
        frmFacDtosAsignar.Show vbModal
    Case id_ActualizarPrecios '= 337                 'Datos Básicos Ventas(3)
         'Actualizar precios actuales y especiales
        frmFacActPrecios2.Proveedor = False
        frmFacActPrecios2.Show vbModal
    Case id_Copiarpreciosdesdecompra '= 338        'Datos Básicos Ventas(3)
          'Copiar desde compra
        CadenaDesdeOtroForm = ""
        AbrirListado2 28
    Case id_InfControlmargenTarifa '= 339          'Datos Básicos Ventas(3)
        'Informe control margenes de tarifas
        AbrirListado (245)
    Case id_CorregirerroresyActTarifas '= 340     'Datos Básicos Ventas(3)
        'Correcion
        AbrirListado 247
    Case id_ControlerrorDtosCliente '= 341         'Datos Básicos Ventas(3)
        AbrirListado3 13
    Case id_TarifasTaxímetros '= 342                 'Datos Básicos Ventas(3)
        frmTaxcoTarifa.Show vbModal
    End Select
End Sub


Private Sub AbrirFormOfertasPedidos(ByRef Accion As Long)

    Select Case Accion
    Case id_Ofertas '= 401                            'Ofertas-Pedidos Ventas(4)
        AbrirOfertas
        
    Case id_GrupodePlantillas ' = 402                'Ofertas-Pedidos Ventas(4)
        'Mantenimiento de Grupos de Plantillas
        frmFacGrupoPlantilla.Show vbModal
    
    Case id_Plantillas '= 403                         'Ofertas-Pedidos Ventas(4)
        'Mantenimiento de Plantillas
        frmFacPlantilla.Show vbModal
    Case id_InfOfertasefectuadas '= 404             'Ofertas-Pedidos Ventas(4)
        'Listado de Ofertas Efectuadas
        AbrirListadoOfer (34) '34: Informe Ofertas Efectuadas

    
    Case id_Pedidos, id_HcoPedidosanulados '405 406               'Ofertas-Pedidos Ventas(4)
       'Mantenimiento de Pedidos  Y Histórico de Pedidos
       
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntPedidos.DatosADevolverBusqueda2 = ""
            frmFacEntPedidos.EsHistorico = Accion = id_HcoPedidosanulados
            frmFacEntPedidos.Show vbModal
        Else
            frmFacEntPedSail.DatosADevolverBusqueda2 = ""
            frmFacEntPedSail.EsHistorico = Accion = id_HcoPedidosanulados
            frmFacEntPedSail.Show vbModal
        End If
    
        
    
    Case id_CartasConfirmaciónPedidos '= 407       'Ofertas-Pedidos Ventas(4)
        AbrirListadoOfer (40)
        
    Case id_InfPedidosxArtículo '= 408             'Ofertas-Pedidos Ventas(4)
         'Informe de Pedidos por Articulo
        AbrirListadoPed (41)
    Case id_InfPedidosxCliente '= 409              'Ofertas-Pedidos Ventas(4)
        'Informe de Pedidos por Cliente
        AbrirListadoPed (44)
    Case id_InfDisponibilidadStocks '= 410          'Ofertas-Pedidos Ventas(4)
        'Resumen de Disponibilidad de Stocks
        AbrirListadoPed (42)
    Case id_ImpresiónPedidosxZona '= 411           'Ofertas-Pedidos Ventas(4)
        frmListado2.Opcion = 26
        frmListado2.Show vbModal
    
    Case id_ConsultaPreciosxCliente '= 412         'Ofertas-Pedidos Ventas(4)
        frmFacConsultaPrecios2.Fecha = Now
        frmFacConsultaPrecios2.Show vbModal
        
    Case id_Devoluciónmaterial '= 413                'Ofertas-Pedidos Ventas(4)
        frmFacEntPedidRMA.Show vbModal
    Case id_InfPedidosxdia '= 414                  'Ofertas-Pedidos Ventas(4)
        frmListado5.OpcionListado = 36
        frmListado5.Show vbModal
    Case id_Alertas '= 415                            'Ofertas-Pedidos Ventas(4)
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            frmMensajes.OpcionMensaje = 26
            frmMensajes.Show vbModal
        Else
            frmRepAvisos.Show vbModal
        End If
    End Select
End Sub



Private Sub AbrirFormAlbaranesVenta(ByRef Accion As Long)
    
    Select Case Accion
    Case id_Albaranes            ' 501                          'Albaranes Ventas(5)
        AbrirAlbaranes "ALV", False
    Case id_HcoAlbaranesanulados            ' 502             'Albaranes Ventas(5)
        AbrirAlbaranes "ALV", True
    Case id_AlbaranesDevolución            ' 503               'Albaranes Ventas(5)
        AbrirAlbaranes "DEV", False
    Case id_InfAlbaranesxArtículo            ' 504           'Albaranes Ventas(5)
        'Informe de Albaranes por Articulo
        AbrirListadoPed (49)
    Case id_InfIncumplimientoentrega            ' 505         'Albaranes Ventas(5)
        'Incumplimiento de los Plazos de Entrega
        AbrirListadoPed (51)
    Case id_InfSituaciónAlbaranes            ' 506            'Albaranes Ventas(5)
         frmListado2.Opcion = 23
        frmListado2.Show vbModal
    Case id_ControlAlbaranes, id_ControlAlbaranesFact  ' 507,508                  'Albaranes Ventas(5)
    
         frmFacFacAsignar.Show vbModal
    
    
    Case id_ControlDirecEnvío            ' 509                'Albaranes Ventas(5)
        frmListado4.Opcion = 11
        frmListado4.Show vbModal
    Case id_ImpresiónAlbTransporte            ' 510           'Albaranes Ventas(5)
         frmListado2.Opcion = 27
         frmListado2.Show vbModal
    Case id_InfAlbaranes            ' 511                      'Albaranes Ventas(5)
        frmListado5.OpcionListado = 24
        frmListado5.Show vbModal
        
    Case id_InfAlbaranesentregados            ' 512           'Albaranes Ventas(5)
        frmAvisosAlb.Show vbModal
    Case id_FacturasMostrador            ' 513                 'Albaranes Ventas(5)
        AbrirAlbaranes "ALM", False
    Case id_FacturasRectificativas            ' 514            'Albaranes Ventas(5)
        AbrirAlbaranes "ALR", False
    Case id_AlbaranesServicios            ' 515                'Albaranes Ventas(5)
        AbrirAlbaranes "ALS", False
    Case id_FacturaciónServicios            ' 516              'Albaranes Ventas(5)
    Case id_AlbaranesInternos            ' 517                 'Albaranes Ventas(5)
        AbrirAlbaranes "ALI", False
    Case id_FacturaciónInternos            ' 518               'Albaranes Ventas(5)
    Case id_InfAlbaranesInternos            ' 519             'Albaranes Ventas(5)
    Case id_AlbaranesGasolinera            ' 520               'Albaranes Ventas(5)
        AbrirAlbaranes "ALV", False
    Case id_AlbaranesTienda            ' 521                   'Albaranes Ventas(5)
    
    Case id_ImportarficheroGasolinera            ' 523        'Albaranes Ventas(5)
    Case id_CambiarAlbaranes_Facturas            ' 524         'Albaranes Ventas(5)
    Case id_Previsión            ' 525                          'Albaranes Ventas(5)
            frmListadoPed.codClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
            AbrirListadoPed (50) 'NO IMPRIME LISTADO
    Case id_Combustible            ' 526                        'Albaranes Ventas(5)
    Case id_Tienda            ' 527                             'Albaranes Ventas(5)
    
    Case id_AjusteFormasdePago            ' 528              'Albaranes Ventas(5)
    
    
    Case id_AlbaranesTelefonía            ' 529                'Albaranes Ventas(5)
         AbrirAlbaranes "ALT", False
         
    Case id_ImportarficheroTelefonía, id_Datospendientesfacturar           ' 530 531         'Albaranes Ventas(5)
           If Accion = id_Datospendientesfacturar Then
                'Importacion
                CadenaDesdeOtroForm = ""
                frmTelefono1.Opcion = 1 'Importar
                frmTelefono1.Show vbModal
                If CadenaDesdeOtroForm = "" Then Exit Sub
                Screen.MousePointer = vbHourglass
                Espera 0.2
                DoEvents
                FuerzaCiereFormTelefonia
            End If
        
            frmTelefono1.Opcion = 0
            frmTelefono1.Show vbModal
    
    Case id_CargosVarios            ' 534                      'Albaranes Ventas(5)
        'cargosvarios
        frmTelBolbaiteGR.QueOpcion = 3
        frmTelBolbaiteGR.Show vbModal
    Case id_Modificaciónmasivacuotas            ' 535         'Albaranes Ventas(5)
            'modifiacion masiva
        frmListado4.Opcion = 10
        frmListado4.Show vbModal

    ' 536 537 538  539
    Case id_Comparativadescuentos, id_Facturaciónporsoporte, id_Resumenporsoporte, id_Datosimportaciónfichero
    
        ' 2.- Listado descuentos comprataiivo copera
        ' 3.- Rsumen fracion
        ' 4.- Datos face
        '
        ' 6.-  Datos importados (index=!4)
        If Accion = id_Datosimportaciónfichero Then
             frmTelefono1.Opcion = 6
        Else
            frmTelefono1.Opcion = Accion - 534 '2,3,4
        End If
        frmTelefono1.Show vbModal
        
    Case id_Conceptosconsumos            ' 540                 'Albaranes Ventas(5)
            'Conceptos consumo
            frmTelBolbaiteGR.QueOpcion = 1
            frmTelBolbaiteGR.Show vbModal

    Case id_Descuentosconsumos            ' 541                'Albaranes Ventas(5)
        frmTelDtoConsumo.Show vbModal
    Case id_Conceptoscuotas            ' 542                   'Albaranes Ventas(5)
        frmTelBolbaiteGR.QueOpcion = 0
        frmTelBolbaiteGR.Show vbModal
    Case id_Descuentoscuotas            ' 543                  'Albaranes Ventas(5)
        frmTelDtoCuotas.Show vbModal
    Case id_Cuotaspropias            ' 544                     'Albaranes Ventas(5)
        frmTelBolbaiteGR.QueOpcion = 2
        frmTelBolbaiteGR.Show vbModal

        
    Case id_Parámetros To id_InfTasaspendcobro
        '****************************************************
        '               AGUA
        '*************************************************
        '      desde 545 hasta 555
        'id_Parámetros  id_Calibres     id_Contadores     id_Importarfichero     id_Facturaciónagua
        'id_Resumenfacturación     id_InfFacturaciónxperiodo    id_InfContadoresexportación
        'id_Modificarcuotavarios     id_Declaracióndetalladaejercicio    id_InfTasaspendcobro
        AbrirFormsAgua Accion
    
        
    Case id_Materiasactivas To id_Ajustecomprastrat
        '****************************************************
        '               TRATAMIENTOS
        '*************************************************
        '      desde 556 hasta 563
         'id_Materiasactivas  id_ADR   id_Plagas   id_Flotas id_Tratamientos
         ' id_Partesdetrabajo id_InfFitosanitarios_Campos id_Ajustecomprastrat
         AbrirFormsTratamiento Accion
        
    Case id_Capítulos To id_Facturaciónderrama
        '****************************************************
        '               OBRAS   y gestion parcelas
        '*************************************************
        '      desde 564 hasta 572
        'id_Ajustecomprastrat id_Capítulos  id_Actuaciones id_Partesdetrabajo2
        'id_Tiposordenes id_Reloj id_InfCompras_Ventasactuación id_ImpresiónCertificación
        'id_InfHuertos_Hanegadas id_Facturaciónderrama
        AbrirFormsObras Accion
    
    Case id_PrevisiónFActuración To id_InfTicketsfacturados
        '****************************************************
        '               FACTURACION
        '*************************************************
        '      desde 601 al 610
        ' id_PrevisiónFActuración  id_Facturaciónalbaranes id_Facturarcliente id_HistóricoFacturasVenta
        ' id_ReimpresiónFacturas id_EnvíoFacturasxmail id_EnvíoFacturasweb id_ContabilizarFacturas
        ' id_Contabilizarticketsagrupados id_InfTicketsfacturados
        AbrirFormsFacturacion Accion


    Case id_InfporCliente To id_InfCostes
        '****************************************************
        '               Estadisticas
        '*************************************************
        ' desde 611 al 624
        ' id_InfporCliente   id_InfporTrabajador id_Infpormeses  id_InfporFamilia_Artículo
        ' id_InfporArtículo  id_InfporProveedor  id_InfporAgente id_DetalledeFacturación
        ' id_Margenventas    id_Infportipoprecio id_Artículosmayorventa id_InfporFamiliaagrupado
        ' id_Infportipopedido id_InfCostes
        AbrirFormsInformesEstadisticas Accion


    End Select

End Sub




























'Accoeiones para abrir algunos forms
Private Sub AbrirOfertas()


  If vParamAplic.TipoFormularioClientes = 0 Then
           
        Debug.Assert False
        If False Then
            
            
               EulerParam = DevuelveDesdeBD(conAri, "pathDocs", "eulerparam", "1", "1")
            
               frmFacEntOfertasGR.DatosOferta = ""
               frmFacEntOfertasGR.Show vbModal
            Else
               frmFacEntOfertas2.DatosOferta = ""
               frmFacEntOfertas2.EsHistorico = False 'Index = 5
               frmFacEntOfertas2.Show vbModal
            End If
        Else
            frmFacEntOferSAIL.DatosOferta = ""
            frmFacEntOferSAIL.EsHistorico = False 'Index = 5
            frmFacEntOferSAIL.Show vbModal
        End If

End Sub


Private Sub AbrirAlbaranes(TipoMOVI As String, HCO As Boolean)
    'Abre el formulario de Albaranes para introducir el Albaran de Mostrador
    'y desde este generar la Factura de mostrador
    If vParamAplic.TipoFormularioClientes = 0 Then
        If vParamAplic.HaciendoFrmulariosGrandes Then
            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranesGR.hcoCodTipoM = TipoMOVI
            frmFacEntAlbaranesGR.EsHistorico = HCO
            frmFacEntAlbaranesGR.Show vbModal

        
        Else
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes2.hcoCodTipoM = TipoMOVI
            frmFacEntAlbaranes2.EsHistorico = HCO
            frmFacEntAlbaranes2.Show vbModal
        End If
        
    Else
            frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbSAIL.hcoCodTipoM = TipoMOVI
            frmFacEntAlbSAIL.EsHistorico = HCO
            frmFacEntAlbSAIL.Show vbModal
    End If
End Sub


Private Sub FacturarAlbaranes()
Dim B As Boolean
    'Facturacion de Albaranes de Ventas
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        MsgBox "No puede facturar desde este punto de menú.", vbExclamation
        Exit Sub
    End If
    
    B = False
    If vParamAplic.TipoFormularioClientes = 0 Then
        B = True
    Else
        If vParamAplic.NumeroInstalacion = vbTaxco Then B = True
    End If
    If B Then

        frmListadoPed.codClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
        
    Else
        'PARA sail
        frmFacturaCliSail.Show vbModal
    End If
End Sub




Private Sub AbrirFormsAgua(ByRef Accion As Long)

    Select Case Accion
    Case id_Parámetros            ' 545                         'Albaranes Ventas(5)
        frmAguaParamGR.Show vbModal
    Case id_Calibres            ' 546                           'Albaranes Ventas(5)
        frmAguaCalibresGR.Show vbModal
    Case id_Contadores            ' 547                         'Albaranes Ventas(5)
        frmAguaContadoresGR.Show vbModal
    Case id_Importarfichero            ' 548                   'Albaranes Ventas(5)
        'importar fichero
        AbrirListado3 52
    Case id_Facturaciónagua            ' 549                   'Albaranes Ventas(5)
        'Facturar
        AbrirListado3 51
    Case id_Resumenfacturación            ' 550                'Albaranes Ventas(5)
           'Resumen facturacion 53
        AbrirListado3 53
    Case id_InfFacturaciónxperiodo            ' 551          'Albaranes Ventas(5)
         
        'Listado para rellenar modelos 100,101,102 EPSAR
        'de facturaciones canon generalitat
        AbrirListado3 55
        
    Case id_InfContadoresexportación            ' 552         'Albaranes Ventas(5)
        
        'Listado exportacion contadores
        AbrirListado3 60
        
    Case id_Modificarcuotavarios            ' 553             'Albaranes Ventas(5)
        'Modificar cuota varios
        frmListado4.Opcion = 13
        frmListado4.Show vbModal
    Case id_Declaracióndetalladaejercicio            ' 554     'Albaranes Ventas(5)
         'Declaracion detallada ejereccio
        AbrirListado3 58
    Case id_InfTasaspendcobro            ' 555               'Albaranes Ventas(5)
        'Tasas pendentes de cobro
        AbrirListado3 74
    End Select
End Sub


Private Sub AbrirFormsTratamiento(ByRef Accion As Long)

    Select Case Accion
    Case id_Materiasactivas            ' 556                   'Albaranes Ventas(5)
        frmAlmMatAct.DatosADevolverBusqueda = ""
        frmAlmMatAct.Show vbModal
    Case id_ADR            ' 557                                'Albaranes Ventas(5)
        frmAlmADR.DatosADevolverBusqueda = ""
        frmAlmADR.Show vbModal
    Case id_Plagas            ' 558                             'Albaranes Ventas(5)
        frmAlmPlagas.DatosADevolverBusqueda = ""
        frmAlmPlagas.Show vbModal
    Case id_Flotas            ' 559                             'Albaranes Ventas(5)
        frmFlotas.Show vbModal
    Case id_Tratamientos            ' 560                       'Albaranes Ventas(5)
         frmADVTratamientos.DatosADevolverBusqueda = False
         frmADVTratamientos.Show vbModal
    Case id_Partesdetrabajo            ' 561                  'Albaranes Ventas(5)
        frmADVTraPartes.Show vbModal
    Case id_InfFitosanitarios_Campos            ' 562          'Albaranes Ventas(5)
         frmListado5.OpcionListado = 12
        frmListado5.Show vbModal
    Case id_Ajustecomprastrat            ' 563               'Albaranes Ventas(5)
        frmListado5.OpcionListado = 9
        frmListado5.Show vbModal
    End Select
End Sub


Private Sub AbrirFormsObras(ByRef Accion As Long)

    Select Case Accion
    Case id_Capítulos            ' 564                          'Albaranes Ventas(5)
         frmObraCapitulo.Show vbModal
    Case id_Actuaciones            ' 565                        'Albaranes Ventas(5)
        frmObraActua.Show vbModal
    Case id_Partesdetrabajo2            ' 566                  'Albaranes Ventas(5)
         If InstalacionEsEulerTaxco Then
            frmEulerTrab.Show vbModal
        Else
            frmObrpartesTra.Show vbModal
        End If
    Case id_Tiposordenes            ' 567                      'Albaranes Ventas(5)
        frmObraOT.Show vbModal
    Case id_Reloj            ' 568                              'Albaranes Ventas(5)
         frmEulerReloj.Show vbModal
    Case id_InfCompras_Ventasactuación            ' 569       'Albaranes Ventas(5)
              'Sept 2012
        frmObraListado.Opcion = 3
        frmObraListado.Show vbModal
    Case id_ImpresiónCertificación            ' 570            'Albaranes Ventas(5)
          'Imprimir certificacion
        frmFacturaCliSail.ImprimirCertificacion = True
        frmFacturaCliSail.Show vbModal
    
    
        'GESTION parcelas
    Case id_InfHuertos_Hanegadas, id_Facturaciónderrama           ' 571 572             'Albaranes Ventas(5)
        If Accion = id_InfHuertos_Hanegadas Then
            frmListado5.OpcionListado = 15
            frmListado5.Show vbModal
        Else
            AbrirListado3 72
        End If
    
    End Select
    
End Sub



Private Sub AbrirFormsFacturacion(ByRef Accion As Long)

    Select Case Accion
    Case id_PrevisiónFActuración             ' 601              'Facturación Ventas(6)
        ' Previsión Facturacion de Albaranes
        frmListadoPed.codClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (50) 'NO IMPRIME LISTADO

    Case id_Facturaciónalbaranes             ' 602              'Facturación Ventas(6)
        FacturarAlbaranes
    
    Case id_Facturarcliente             ' 603                   'Facturación Ventas(6)
        
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacturacionCli.Show vbModal
        Else
            frmFacturaCliSail.ImprimirCertificacion = False
            frmFacturaCliSail.Show vbModal
        End If

    Case id_HistóricoFacturasVenta             ' 604           'Facturación Ventas(6)
            
            frmFacHcoFacturas2.hcoCodMovim = ""
            frmFacHcoFacturas2.Show vbModal
            
    
    Case id_ReimpresiónFacturas             ' 605               'Facturación Ventas(6)
        'Reimprimir Factuas ya contabilizadas
        AbrirListadoOfer 226
    
    Case id_EnvíoFacturasxmail, id_EnvíoFacturasweb   ' 606,607              'Facturación Ventas(6)
        AbrirListadoOfer IIf(Accion = id_EnvíoFacturasxmail, 315, 316)
    
    Case id_ContabilizarFacturas             ' 608              'Facturación Ventas(6)
        'Contabilizar Facturas
        AbrirListado (223) 'Para pedir datos
    Case id_Contabilizarticketsagrupados, id_InfTicketsfacturados  ' 609    610             'Facturación Ventas(6)
        AbrirListado2 IIf(Accion = id_Contabilizarticketsagrupados, 12, 13)
        
    End Select
End Sub

Private Sub AbrirFormsInformesEstadisticas(ByRef Accion As Long)

    Select Case Accion
    Case id_InfporCliente             ' 611                   'Facturación Ventas(6)
        'Estadistica Ventas por cliente
        AbrirListadoPed (227)
        BorrarTempInformes
    Case id_InfporTrabajador             ' 612                'Facturación Ventas(6)
        'Estadistica Ventas por Trabajador
        AbrirListadoPed (228)
    Case id_Infpormeses             ' 613                     'Facturación Ventas(6)
        'Estadistica Ventas por Meses
        AbrirListadoPed (229)
    Case id_InfporFamilia_Artículo             ' 614          'Facturación Ventas(6)
        'Listado de estadistica ventas por familia de articulo
        AbrirListadoOfer (230)
    Case id_InfporArtículo             ' 615                  'Facturación Ventas(6)
        AbrirListado3 18
    Case id_InfporProveedor             ' 616                 'Facturación Ventas(6)
        'Por proveedor
        AbrirListado2 6
    Case id_InfporAgente             ' 617                    'Facturación Ventas(6)
        'Ventas por agente
        AbrirListado2 16
    Case id_DetalledeFacturación             ' 618             'Facturación Ventas(6)
        'Detalle facturacion clientes
        AbrirListadoOfer (231)
    Case id_Margenventas             ' 619                      'Facturación Ventas(6)
        'Estadistica margen ventas por artículo
        AbrirListado (246)
    Case id_Infportipoprecio             ' 620               'Facturación Ventas(6)
        'Vtas x tipo d precio
        AbrirListado3 38
    Case id_Artículosmayorventa             ' 621              'Facturación Ventas(6)
        'Articulos ams vendidos
        AbrirListado3 39
    Case id_InfporFamiliaagrupado             ' 622          'Facturación Ventas(6)
        frmListado4.Opcion = 7
        frmListado4.Show vbModal
    Case id_Infportipopedido             ' 623               'Facturación Ventas(6)
        'Listado pedidos por "peticion " cliente (si-no)
        AbrirListado3 63
    Case id_InfCostes             ' 624                        'Facturación Ventas(6)
        AbrirListado2 53  'Costes euler

    End Select
End Sub



Private Sub AbrirFormsCompras(ByRef Accion As Long)

    Select Case Accion
    Case id_Proveedores             ' 901                        'Compras(9)
        frmComProveedoresGr.Show vbModal
        
    Case id_ProveedoresVarios             ' 902                 'Compras(9)
        frmComProveV.Show vbModal
    Case id_DireccionesCompra             ' 903                 'Compras(9)
    
    Case id_EtiquetasdeProveedores             ' 904           'Compras(9)
        AbrirListadoOfer (305) '305: Informe Etiquetas de Proveedores
    
    Case id_CartasaProveedores             ' 905               'Compras(9)
        AbrirListadoOfer (306) '306: Informe Cartas a Proveedores
       
    Case id_EtiquetasdebultosCompra             ' 906                'Compras(9)
        AbrirListado 101 'bultos compra
    Case id_PreciosProveedor             ' 907                  'Compras(9)
         'precios proveedor
        frmComPreciosProv2.NuevoDato = "" 'Para que no se poing en modo insercion
        frmComPreciosProv2.Show vbModal
    Case id_DescuentosProveedor             ' 908               'Compras(9)
        'Dto proveedor
        frmComDtosFamMarca.Show vbModal
    Case id_CopiarPreciosdesdeventa             ' 909         'Compras(9)
        'Copiar desde venta
        CadenaDesdeOtroForm = "V"
        AbrirListado2 28
    Case id_ActualizarPreciosCompra             ' 910          'Compras(9)
        frmFacActPrecios2.Proveedor = True
        frmFacActPrecios2.Show vbModal
    Case id_PedidosProveedor, id_HcoPedidosanuladosCompra      ' 912 911                  'Compras(9)
    
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmComEntPedidos2.MostrarDatos = ""
            frmComEntPedidos2.EsHistorico = Accion = id_HcoPedidosanuladosCompra
            frmComEntPedidos2.Show vbModal
        Else
            'SAIL
            frmComEntPedidosSa.MostrarDatos = ""
            frmComEntPedidosSa.EsHistorico = Accion = id_HcoPedidosanuladosCompra
            frmComEntPedidosSa.Show vbModal
        End If
        
    Case id_InfMaterialpdterecibir             ' 913        'Compras(9)
        AbrirListadoOfer (307) '307: List. Materia pte recibir
    Case id_PropuestadePedido             ' 914                'Compras(9)
         AbrirListado2 32
    Case id_InfReaprovisionamiento             ' 915           'Compras(9)
        AbrirListado3 71
        
    Case id_AlbaranesProveedor, id_HcoAlbaranesanuladosCompra ' 917 916
    
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmComEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmComEntAlbaranesGR.EsHistorico = Accion = id_HcoAlbaranesanuladosCompra
            frmComEntAlbaranesGR.Show vbModal
         Else
            frmComEntAlbaranSA.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmComEntAlbaranSA.EsHistorico = Accion = id_HcoAlbaranesanuladosCompra
            frmComEntAlbaranSA.Show vbModal
        
        End If
    Case id_InfPendientefacturar             ' 918            'Compras(9)
        'Listado de Albaranes pendientes de Factura
        AbrirListadoOfer (308) '308: List. Albaranes pte facturar

    Case id_ControlAlbaranesCompra             ' 919                  'Compras(9)
        frmComCtrDoc.Show vbModal
    Case id_ControlAlbaranesfacturados             ' 920       'Compras(9)
        frmComCtrDoc.Show vbModal
    Case id_RecepciónFacturas             ' 921                 'Compras(9)
        frmComFacturarGR.Codprove = -1
        frmComFacturarGR.Show vbModal
    Case id_HistóricoFacturasCompra             ' 922          'Compras(9)
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmComHcoFacturas2GR.hcoCodMovim = ""
            frmComHcoFacturas2GR.Show vbModal
        Else
            'SAIL
            frmComHcoFacturSA.hcoCodMovim = ""
            frmComHcoFacturSA.Show vbModal
        End If
        
    Case id_ContabilizarFacturasCompra             ' 923              'Compras(9)
        'Contabilizar Facturas
        AbrirListado (224) 'Para pedir datos
    
    Case id_InfporProveedorCompra             ' 924                 'Compras(9)
         'Listado de compras por proveedor
        AbrirListadoOfer (310)
    Case id_InfporFamilia_ArtículoCompra             ' 925          'Compras(9)
         'Listado de compras por Familia
        AbrirListadoOfer (311)
    Case id_InfpormesesCompra             ' 926                     'Compras(9)
        frmVarios.Opcion = 11
        frmVarios.Show vbModal
    Case id_InfAlbaranesxProveedor             ' 927         'Compras(9)
        'Listado de alb compras por proveedor
        AbrirListadoOfer (312)
    Case id_InfPrevisiónpagos             ' 928               'Compras(9)
         'frmListado3.Opcion = 7
         'frmListado3.Show vbModal
         AbrirListado3 7
    Case id_InfProveedor_Marca_Familia             ' 929       'Compras(9)
        AbrirListado2 50  'compras familia-marca
    End Select


End Sub


Private Sub AbrirFormsAministracion(ByRef Accion As Long)
    
    Select Case Accion
    Case id_Trabajadores             ' 1001                       'Administración(10)
        If vUsu.Nivel2 = 2 Then Exit Sub
        frmAdmTrabajadores.Show vbModal
    Case id_GastosTécnicos             ' 1002                    'Administración(10)
        frmAdmGasTec.Show vbModal
    Case id_NóminasyGastos             ' 1003                   'Administración(10)
         frmAdmNominas.Show vbModal
    Case id_CálculoRiesgo             ' 1004                     'Administración(10)
        If vUsu.Nivel > 0 Then
            MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
        Else
            AbrirListado2 31
        End If
        
    Case id_InfRiesgo             ' 1005                        'Administración(10)
        frmInformesNew6.OpcionListado = 1
        frmInformesNew6.Show vbModal
        
    Case id_Correccióncostesvariosfactura             ' 1007     'Administración(10)
    
    Case id_CorreccióncostesEstVentas             ' 1008       'Administración(10)
        'Modificar coste estadistica ventas
        AbrirListado3 11
    Case id_InfVentasacrédito             ' 1009              'Administración(10)
        AbrirListado3 25
    Case id_BeneficioProveedor             ' 1010                'Administración(10)
        'beneficio por proveedor
        AbrirListado2 40
    Case id_BeneficioCliente             ' 1011                  'Administración(10)
        AbrirListado2 41
    Case id_BeneficioMarca_Agente_Proveedor             ' 1012     'Administración(10)
        'Beneficio marca-agente-proveedor
        AbrirListado2 48
    Case id_InfArtículosenpromoción             ' 1013        'Administración(10)
        AbrirListado3 5
    Case id_InfArtículosconDtoEspecial             ' 1014     'Administración(10)
        AbrirListado3 34
    Case id_InfVentasTrabajadordía             ' 1015         'Administración(10)
        'Ventas trabajaodr x dia
        AbrirListado3 9
    Case id_InfVentasxFPago             ' 1016               'Administración(10)
        'Listado ventas por forma de pago
        AbrirListado3 19
    Case id_InfComparativoDtosCompra_Venta             ' 1017     'Administración(10)
        frmListado5.OpcionListado = 37
        frmListado5.Show vbModal
        
    Case id_ResumenVentasAgente             ' 1018              'Administración(10)
        'Ventas x agente
        AbrirListado2 36
    Case id_BeneficioAgente             ' 1019                   'Administración(10)
        'beneficio por agente
        AbrirListado2 37
    Case id_InfVentasAgente_Trabajador             ' 1020      'Administración(10)
        AbrirListado3 37
    Case id_InfComisionesECO             ' 1021                'Administración(10)
        frmFacComisionAgen.Show vbModal
    Case id_InfAgente_Familia_Marca             ' 1022          'Administración(10)
        AbrirListado3 46
    Case id_InfAgente_Marca_Familia             ' 1023          'Administración(10)
        AbrirListado2 49
    Case id_RegistroGastos             ' 1024                    'Administración(10)
        'regeistros flota
        frmFlotaReg.DatosADevolverBusqueda = ""
        frmFlotaReg.Show vbModal
    Case id_FlotasAdm             ' 1025                             'Administración(10)
        frmFlotas.Show vbModal
    Case id_Conceptos             ' 1026                          'Administración(10)
        frmFlotasConceptos.DatosADevolverBusqueda = ""
        frmFlotasConceptos.Show vbModal
    End Select

End Sub



Private Sub AbrirFormsMantenimiento(ByRef Accion As Long)

    Select Case Accion
    Case id_TiposdeContrato                     ' 1101                  'Mantenimientos(11)
        frmManTiposContrato.Show vbModal
    Case id_Mantenimientos                     ' 1102                     'Mantenimientos(11)
        frmManMantenimientosGR.Show vbModal
    Case id_InfMantenimientos                     ' 1103                'Mantenimientos(11)
        'Listados de Mantenimientos
        AbrirListado 70
    Case id_InfRevisiones                     ' 1104                    'Mantenimientos(11)
        'Listado Revisiones de Mantenimientos
        AbrirListado 71
    Case id_FichaMantenimientos                     ' 1105               'Mantenimientos(11)
        'Listado Fichas de Mantenimientos
        AbrirListado 72
    Case id_InfAltas                     ' 1106                         'Mantenimientos(11)
        'Listado Altas de Mantenimientos
        AbrirListado 73
    Case id_InfTeórico                     ' 1107                       'Mantenimientos(11)
        AbrirListado 77
    Case id_Etiquetas                     ' 1108                          'Mantenimientos(11)
        AbrirListado 79
    Case id_Cartasrenovación                     ' 1109                  'Mantenimientos(11)
        AbrirListado 78
    Case id_Traspasosiguienteaactual                     ' 1110        'Mantenimientos(11)
        frmMensajes.OpcionMensaje = 18
        frmMensajes.Show vbModal
    Case id_HcoMantenimientos                     ' 1111                'Mantenimientos(11)
         frmManMantenimientosAnuGR.Show vbModal
    Case id_InfAnulados                     ' 1112                      'Mantenimientos(11)
        AbrirListado 76
        
    Case id_PrevisiónMantenimientos                     ' 1113           'Mantenimientos(11)
        AbrirListadoPed (74) 'NO IMPRIME LISTADO
    Case id_FacturaciónMantenimientos                     ' 1114         'Mantenimientos(11)
        AbrirListadoPed (75) 'NO IMPRIME LISTADO
    Case id_PrevisiónRenting                     ' 1115                  'Mantenimientos(11)
        frmListado3.OtrosDatos = ""
        AbrirListado3 23
    Case id_FacturaciónRenting                     ' 1116                'Mantenimientos(11)
        frmListado3.OtrosDatos = ""
        AbrirListado3 22
    End Select
End Sub



Private Sub AbrirFormsReparaciones(ByRef Accion As Long)


    Select Case Accion
    Case id_NúmerosdeSerie                     ' 1201                   'Reparaciones(12)
        'Mantenimiento de Nºs de Serie
        frmRepNumSerie2GR.Show vbModal
    Case id_Motivosbaja                     ' 1202                       'Reparaciones(12)
        'Motivos baja equipos
        frmRepMotivosBaja.Show vbModal
    Case id_MotivosPendienteRep                     ' 1203             'Reparaciones(12)
        'Motivos Pendientes Reparar
        frmRepMotivosPend.Show vbModal
    Case id_Tiposaveria                     ' 1204                       'Reparaciones(12)
        frmtipave.Show vbModal
    Case id_Trabajosrealizados                     ' 1205                'Reparaciones(12)
        frmManTraReali.Show vbModal
        
    Case id_Serviciosasistenciatécnica                     ' 1206       'Reparaciones(12)
        frmManSat.Show vbModal
    Case id_EntradaReparación, id_ControlReparación, id_HcoReparaciones  ' 1210  1211 1212

        frmRepEntReparacionesGR.EntradaEquipo = ""
        frmRepEntReparacionesGR.ControlRep = IIf(Accion = id_ControlReparación, True, False)
        frmRepEntReparacionesGR.EsHistorico = Accion = id_HcoReparaciones
        frmRepEntReparacionesGR.Show vbModal
    
    
    Case id_Infpordía                     ' 1213                       'Reparaciones(12)
        'Listado de las Reparaciones del dia
        AbrirListado (63)
    Case id_InfporClienteRepar                     ' 1214                   'Reparaciones(12)
        'Listado de las Reparaciones por cliente
        AbrirListado (64)
    Case id_FrecuenciadeReparación                     ' 1215           'Reparaciones(12)
        'Listado de Frecuencia de Reparaciones
        AbrirListado (406)
    Case id_InfporTécnico                     ' 1216                   'Reparaciones(12)
        AbrirListado2 2
    Case id_InfReparacionesefectuadas                     ' 1217       'Reparaciones(12)
        AbrirListado2 1
    Case id_InfGarantíaproveedor                     ' 1218            'Reparaciones(12)
        AbrirListado2 30
        
        
    Case id_AlbaránOrdendetrabajo, id_AlbaránReparación, id_AlbaránServExterior  ' 1219 1220 1221
    
    
            'Alkbaranes.   ALE ALO ALR
            If Accion = id_AlbaránOrdendetrabajo Then
                frmFacEntAlbSAIL.hcoCodTipoM = "ALO"
            ElseIf Accion = id_AlbaránServExterior Then
                frmFacEntAlbSAIL.hcoCodTipoM = "ALE"
            Else
                frmFacEntAlbSAIL.hcoCodTipoM = "ALR"
            End If
            frmFacEntAlbSAIL.EsHistorico = False
            frmFacEntAlbSAIL.Show vbModal
    
    
    Case id_Proyectos                     ' 1222                          'Reparaciones(12)
        'Proyectos
        frmFacProyecto.Show vbModal
        
    Case id_Frecuencias                     ' 1207                        'Reparaciones(12)
        frmFrecuenciasGR.Show vbModal
    Case id_AlbaránReparaciónRepara                     ' 1223                 'Reparaciones(12)
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALR"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
        
    Case id_PrevisiónFActuraciónRepara                     ' 1224              'Reparaciones(12)
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            frmListadoPed.codClien = "ALO" 'utilizamos esta vble para pasarle el tipo de movimiento
        Else
            frmListadoPed.codClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
        End If
        AbrirListadoPed (50) 'NO IMPRIME LISTADO
    
    Case id_Facturación                     ' 1225                        'Reparaciones(12)
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            frmListadoPed.codClien = "ALO" 'utilizamos esta vble para pasarle el tipo de movimiento
        Else
            frmListadoPed.codClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
        End If
        AbrirListadoPed (52)
    End Select

End Sub



Private Sub AbrirFormsCRM(ByRef Accion As Long)

    Select Case Accion
    Case id_TiposAcciones                     ' 1301                     'CRM(13)
         frmCRMtipos.Show vbModal
    Case id_Conceptosllamadas                     ' 1302                 'CRM(13)
        frmLlamadasTipo.Show vbModal
    Case id_Accionescomerciales                     ' 1303               'CRM(13)
        frmCRMMto.DesdeElCliente = 0 'No clien
        frmCRMMto.TipoPredefinido = 0   'Ninguno
        frmCRMMto.Show vbModal
    Case id_Generaracciones                     ' 1304                   'CRM(13)
        frmCRMVarios.Opcion = 0
        frmCRMVarios.Show vbModal
    Case id_InfMasivo                     ' 1305                        'CRM(13)
        frmListadoOfer.OpcionListado = 406
        frmListadoOfer.Show vbModal
    Case id_InfresumenCRM                     ' 1306                    'CRM(13)
        frmCRMVarios.Opcion = 1
        frmCRMVarios.Show vbModal
    Case id_InfClientesporacción                     ' 1307             'CRM(13)
        frmListado5.OpcionListado = 22
        frmListado5.Show vbModal
    Case id_AvisosdeClientes                     ' 1308                 'CRM(13)
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            frmMensajes.OpcionMensaje = 26
            frmMensajes.Show vbModal
        Else
            frmRepAvisos.Show vbModal
        End If
    Case id_InfAvisospendientes                     ' 1309             'CRM(13)
        'Listado de avisos de averias de clientes pendientes
        AbrirListado (409)
    Case id_Borreavisoscerrados                     ' 1310              'CRM(13)
        AbrirListado 83
    Case id_LlamadasClientes                     ' 1311                  'CRM(13)
        frmLlamadas.Show vbModal
    End Select
End Sub



Private Sub AbrirFormsProduccion(ByRef Accion As Long)

    Select Case Accion
    Case id_Ordenesproducción                     ' 1401                 'Producción(14)
        frmProdOrden.DatosADevolverBusqueda = ""
        frmProdOrden.Show vbModal
    Case id_Ordenesenvasado                     ' 1402                   'Producción(14)
         frmProdEnvas.DatosADevolverBusqueda = ""
        frmProdEnvas.Show vbModal
    Case id_CostesTasas                     ' 1403                       'Producción(14)
        frmAlmDescCostesTasas.Show vbModal
    Case id_Registrotrazabilidad                     ' 1404              'Producción(14)
        frmListLotes.Show vbModal
    Case id_Parámetroscalidad                     ' 1405                 'Producción(14)
         frmAlmCalidad.Show vbModal
    Case id_DeclaraciónalcoholAEAT                     ' 1406           'Producción(14)
        If vParamAplic.NumeroInstalacion = vbFontenas Then
            frmListado5.OpcionListado = 20
            frmListado5.Show vbModal
        End If
    End Select

End Sub


Private Sub AbrirFormsTPV(ByRef Accion As Long)

    Select Case Accion
    Case id_PantalladeVenta                     ' 1501                  'Punto de Venta(15)
        AbirTPVpantallaVenta
    Case id_Cierredecaja                     ' 1502                     'Punto de Venta(15)
        AbrirListadoOfer (240)
    Case id_Etiquetasestantería                     ' 1503               'Punto de Venta(15)
        AbrirListado 94
    Case id_ParámetrosGenerales                     ' 1504               'Punto de Venta(15)
        'Parámetros generales del TPV
        frmFacTPVParamG.Show vbModal
    Case id_ParámetrosTerminales                     ' 1505              'Punto de Venta(15)
        frmFacTPVParamT.Show vbModal
    End Select
End Sub


Private Sub AbrirFormsUtilidades(ByRef Accion As Long)

    Select Case Accion
    Case id_Accionesrealizadas                     ' 2001                'Utilidades(20)
           
            Load frmLog
            DoEvents
            frmLog.Show vbModal
           
    Case id_BorreFacturasyMovimientos                     ' 2002       'Utilidades(20)
        AbrirListado 97
    Case id_CambiodeCliente                     ' 2003                  'Utilidades(20)
        frmListado5.OpcionListado = 35
        frmListado5.OtrosDatos = ""
        frmListado5.Show vbModal
    Case id_InfAlb_Pedanulados                     ' 2004              'Utilidades(20)
        AbrirListado3 16
    Case id_ComprobarCCC_NIF                     ' 2005                  'Utilidades(20)
        'Comprobar cuenta banco secciones(y contabilidades)
        frmListadoOfer.OpcionListado = 408
        frmListadoOfer.Show vbModal
    Case id_ExportarAlbaranesservicio                     ' 2006        'Utilidades(20)
        frmListado3.Opcion = 59
        frmListado3.Show vbModal
    Case id_ExportarmailCSV                     ' 2007                  'Utilidades(20)
        frmListado3.Opcion = 66
        frmListado3.Show vbModal
    
    Case id_Lotesfitossubvencionados, id_DeclaracionROPO   ' 2008-2009        'Utilidades(20)
    
        If Accion = id_Lotesfitossubvencionados Then
            frmFacLotesGeneralitat.Show vbModal
        Else
            frmUtDeclara.Show vbModal
        End If
    
    End Select

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
        MsgBox "No se puedo establecer un terminal.", vbExclamation

    End If
End Sub




Private Sub FuerzaCiereFormTelefonia()
    On Error Resume Next
    Unload frmTelefono1
    If Err.Number <> 0 Then Err.Clear
End Sub



