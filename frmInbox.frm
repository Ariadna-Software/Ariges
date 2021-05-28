VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#17.2#0"; "Codejock.Calendar.v17.2.0.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Begin VB.Form frmInbox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13110
   Icon            =   "frmInbox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ScrollBar ScrollBarCalendar 
      Height          =   5295
      Left            =   8400
      TabIndex        =   2
      Top             =   1080
      Width           =   255
      _Version        =   1114114
      _ExtentX        =   450
      _ExtentY        =   0
      _StockProps     =   64
      Orientation     =   1
      UseVisualStyle  =   0   'False
      Appearance      =   6
   End
   Begin XtremeCalendarControl.CalendarControl CalendarControl 
      Height          =   6135
      Left            =   6720
      TabIndex        =   1
      Top             =   240
      Width           =   8415
      _Version        =   1114114
      _ExtentX        =   14843
      _ExtentY        =   10821
      _StockProps     =   64
      ViewType        =   1
      ShowCaptionBar  =   -1  'True
   End
   Begin VB.Frame FrameBorder 
      BackColor       =   &H000F0200&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.Image Image1 
         Height          =   810
         Left            =   4080
         Picture         =   "frmInbox.frx":058A
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   780
      End
      Begin VB.Image Image2 
         Height          =   19200
         Left            =   0
         Picture         =   "frmInbox.frx":1420
         Top             =   0
         Width           =   28800
      End
   End
End
Attribute VB_Name = "frmInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SORT_BY_COLUMN As Integer
Dim GROUP_BY_COLUMN As Boolean
Dim REMOVE_COLUMN As Boolean
Dim ALIGN_COLUMN As XTPColumnAlignment

'Array used to exctract icons from a bitmap (bitmap in Icons folder)
Dim iconArray(0 To 9) As Long

Dim fntStrike As StdFont
Dim fntBold As StdFont
Dim Profile As String

'calendar defines
Dim ContextEvent As CalendarEvent
Public DisableDragging_ForRecurrenceEvents As Boolean
Public DisableInPlaceCreateEvents_ForSaSu As Boolean
Public EnableScrollV_DayView As Boolean
Public EnableScrollH_DayView As Boolean
Public EnableScrollV_WeekView As Boolean
Public EnableScrollV_MonthView As Boolean
Private nToolTips_Mode As Long


Dim m_pCustomDataHandler As Object

Private Sub CalendarControl_OnReminders(ByVal Action As XtremeCalendarControl.CalendarRemindersAction, ByVal Reminder As XtremeCalendarControl.CalendarReminder)
       
On Error Resume Next

    frmReminders.OnReminders Action, Reminder
    
    If Action = xtpCalendarRemindersFire Then
        frmReminders.Show vbModal, Me
        If Err.Number <> 0 Then Err.Clear
               
        
        
        UltimaLecturaReminders = Timer
    End If
    
End Sub




'Public frmPreviewPane As frmPreviewPane

Private Sub Form_Load()
    
    CargaFondo
    CreateCalendar
        
    
    CalendarControl.SetScrollBars 0, ScrollBarCalendar.hwnd
    
End Sub

Private Sub CargaFondo()
Dim LaImg As String
    On Error GoTo eCargaFondo
    
    
    LaImg = App.Path & "\styles\fondo-ariges II.jpg"
    If Dir(LaImg, vbArchive) = "" Then LaImg = App.Path & "\arifon2.dll"
    
    Image2.Picture = LoadPicture(LaImg)
    
    
    Exit Sub
eCargaFondo:
    Err.Clear
End Sub


Private Sub CreateCalendar()

    CalendarControl.ShowCaptionBarSwitchViewButtons = False
    
    DisableDragging_ForRecurrenceEvents = False
    DisableInPlaceCreateEvents_ForSaSu = False
    EnableScrollV_DayView = True
    EnableScrollH_DayView = True
    EnableScrollV_WeekView = True
    EnableScrollV_MonthView = True
         



    OpenProvider  'cjCalendarData_MySQL, C

    CalendarControl.Populate
    
    CalendarControl.Options.WorkWeekMask = xtpCalendarDayMo_Fr
    CalendarControl.Options.FirstDayOfTheWeek = 1
    
    Dim Today As Date
    Today = Now
    CalendarControl.ViewType = xtpCalendarDayView
    'CalendarControl.DayView.ShowDays Today - 2, Today + 2
    Dim DateMin As Date, DateMax As Date
    DateMin = DateAdd("m", -1, Today)
    DateMin = DateAdd("m", 3, Today)

    
    CalendarControl.DayView.ShowDays DateMin, DateMax
       
    CalendarControl.DayView.ScrollToWorkDayBegin
    
    CalendarControl.Options.DayViewScale2Visible = False
    CalendarControl.Options.DayViewScaleLabel = ""
    
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    
    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
    
    FrameBorder.Move 0, 4, ScaleWidth - 6, ScaleHeight - 10
   
    Image1.Left = Me.Width - 900
    Image1.top = Me.Height - 1000
    If CalendarControl.visible Then
        CalendarControl.Move 1, 4, ScaleWidth - ScrollBarCalendar.Width - 7, ScaleHeight - 11
        
        ScrollBarCalendar.Move CalendarControl.Width, CalendarControl.top + 44, ScrollBarCalendar.Width, CalendarControl.Height - 44

    End If

End Sub

Public Sub SetColor(Id As Integer)

   
   
    Dim HexColor As Long
    If Id = ID_OPTIONS_STYLEBLACK2010 Then
        'HexColor = &H393839
        HexColor = &H949294
    ElseIf Id = ID_OPTIONS_STYLESILVER2010 Then
        'HexColor = &H73716B
        HexColor = &HBDB2AD
    Else
        HexColor = &HBD9E84
    End If
    
    FrameBorder.BackColor = HexColor
 
End Sub


'Subroutine that randomly generates a date.  If you group by a column with a date, the records will
'be grouped by how recent the date is in respect to the current date
Public Sub GetDate(ByVal Item As ReportRecordItem, Optional Week = -1, Optional Day = -1, Optional Month = -1, Optional Year = -1, _
                        Optional Hour = -1, Optional Minute = -1)
    Dim WeekDay As String
    Dim TimeOfDay As String
    
    'Initialize random number generator
    Randomize
        
    'Random number fomula
    'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    
    'If no week day was provided, randomly select a week day
    If (Week = -1) Then
        Week = Int((7 * Rnd) + 1)
    End If
    
    'Determine the week text
    Select Case Week
        Case 1:
            WeekDay = "Sun"
        Case 2:
            WeekDay = "Mon"
        Case 3:
            WeekDay = "Tue"
        Case 4:
            WeekDay = "Wed"
        Case 5:
            WeekDay = "Thu"
        Case 6:
            WeekDay = "Fri"
        Case 7:
            WeekDay = "Sat"
    End Select
    
    'If no month was provided, randomly select a month
    If (Month = -1) Then
        Month = Int((DatePart("m", Now) - 1 + 1) * Rnd + 1)
    End If
     
    'If no day was provided, randomly select a day
    If (Day = -1) Then
        Day = Int((31 - 1 + 1) * Rnd + 1)
    End If
    
    'If no year was provided, randomly select a year
    If (Year = -1) Then
        Year = Int((2004 - 2003 + 1) * Rnd + 2003)
    End If
    
    'If no hour was provided, randomly select a hour
    If (Hour = -1) Then
        Hour = Int((12 - 1 + 1) * Rnd + 1)
    End If
    
    'If no minute was provided, randomly select a minute
    If (Minute = -1) Then
        Minute = Int((60 - 10 + 1) * Rnd + 10)
    End If
     
    'Randomly select AM or PM
    If (Int(2 * Rnd + 1) = 1) Then
        TimeOfDay = "PM"
    Else
        TimeOfDay = "AM"
    End If
       
    'This block of code determines the GroupPriority, GroupCaption, and SortPriority of the Item
    'based on the date or generated provided.  If the date is the current date, then this Item will
    'have High group and sort priority, and will be given the "Date: Today" groupcaption.
    If (Month = DatePart("m", Now)) And (Day = DatePart("d", Now)) And (Year = DatePart("yyyy", Now)) Then
        Item.GroupCaption = "Date: Today"
        Item.GroupPriority = 0
        Item.SortPriority = 0
    ElseIf (Month = DatePart("m", Now)) And (Year = DatePart("yyyy", Now)) Then
        Item.GroupCaption = "Date: This Month"
        Item.GroupPriority = 1
        Item.SortPriority = 1
    ElseIf (Year = DatePart("yyyy", Now)) Then
        Item.GroupCaption = "Date: This Year"
        Item.GroupPriority = 2
        Item.SortPriority = 2
    Else
        Item.GroupCaption = "Date: Older"
        Item.GroupPriority = 3
        Item.SortPriority = 3
    End If
    
    'Assign the DateTime string to the value of the ReportRecordItem
    Item.Value = WeekDay & " " & Month & "/" & Day & "/" & Year & " " & Hour & ":" & Minute & " " & TimeOfDay
End Sub




'***************************************
'Calendar Code
'***************************************

Public Property Get ToolTips_Mode() As Long
   ToolTips_Mode = nToolTips_Mode
End Property

Public Property Let ToolTips_Mode(ByRef nMode As Long)

   nToolTips_Mode = nMode
   CalendarControl.EnableToolTips (nMode = 0)
End Property

Private Sub CalendarControl_ContextMenu(ByVal X As Single, ByVal Y As Single)

    Debug.Print "On context menu"
       
    'Exit Sub
       
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    Dim ChildControl As CommandBarControl
    
    Set Popup = frmppalN.CommandBars.Add("Popup", xtpBarPopup)
        
        Dim HitTest As CalendarHitTestInfo
        Set HitTest = CalendarControl.ActiveView.HitTest
        
        If Not HitTest.ViewEvent Is Nothing Then
            Set ContextEvent = HitTest.ViewEvent.Event
            Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_OPEN, "&Open")
            Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_DELETE, "&Delete")
            Popup.ShowPopup
            Set ContextEvent = Nothing
        ElseIf (HitTest.HitCode = xtpCalendarHitTestDayViewTimeScale) Then
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_NEW, "&New Event")
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_CHANGE_TIMEZONE, "Change Time Zone")
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_60, "6&0 Minutes")
           Control.BeginGroup = True
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_30, "&30 Minutes")
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_15, "&15 Minutes")
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_10, "10 &Minutes")
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_5, "&5 Minutes")
           Popup.ShowPopup
        Else
           Set Control = Popup.Controls.Add(xtpControlButton, ID_CALENDAREVENT_NEW, "&Nuevo evento")
           Popup.ShowPopup
        End If
End Sub

Private Sub CalendarControl_DblClick()

    'Exit Sub


    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl.ActiveView.HitTest
    
    Dim Events As CalendarEvents
    If Not HitTest.HitCode = xtpCalendarHitTestUnknown Then
     '   Set Events = CalendarControl.DataProvider.RetrieveDayEvents(HitTest.ViewDay.Date)
    End If
    
    If HitTest.ViewEvent Is Nothing Then
        mnuNewEvent
    Else
        ModifyEvent HitTest.ViewEvent.Event, False
    End If
End Sub

Private Sub CalendarControl_EventChanged(ByVal EventID As Long)
    Dim pEvent As CalendarEvent
    Set pEvent = CalendarControl.DataProvider.GetEvent(EventID)
    
    ' pEvent Is Nothing for recurrence Ocurrence and Exception.
    ' See CalendarControl_PatternChanged for recurrence events changes.
    If Not pEvent Is Nothing Then
        
    End If
End Sub

Private Sub CalendarControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Debug.Print "KeyDown"
    Dim BeginSelection As Date, EndSelection As Date, AllDay As Boolean

    If CalendarControl.ActiveView.getSelection(BeginSelection, EndSelection, AllDay) Then
        Debug.Print "Selection: "; BeginSelection; " - "; EndSelection
    End If

End Sub

Private Sub CalendarControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl.ActiveView.HitTest
    
    If (Not HitTest.ViewEvent Is Nothing) Then
      ' Debug.Print "MouseMove. HitTest = "; HitTest.ViewEvent.Event.Subject
       
       If ToolTips_Mode = 1 Then
           CalendarControl.ToolTipText = "[" & HitTest.ViewEvent.Event.Id & "]  " & HitTest.ViewEvent.Event.Subject
       Else
           CalendarControl.ToolTipText = ""
           Me.Refresh
       End If
    Else
        CalendarControl.ToolTipText = ""
        Me.Refresh
    End If
    
    'If HitTest.HitCode And xtpCalendarHitTestDayArea Then
    '    Debug.Print "HitTest DayArea"
    'End If
    'If HitTest.HitCode And xtpCalendarHitTestDayHeader Then
    '    Debug.Print "HitTest DayHEADER"
    'End If
    'If HitTest.HitCode And xtpCalendarHitTestDayViewAllDayEvent Then
    '    Debug.Print "HitTest AllDayEvent"
    'End If
    'If HitTest.HitCode And xtpCalendarHitTestDayViewCell Then
    '    Debug.Print "HitTest DayViewCell"
    'End If
    'If HitTest.HitCode And xtpCalendarHitTestDayViewTimeScale Then
    '    Debug.Print "HitTest TimeScale"
    'End If
End Sub

Private Sub CalendarControl_PrePopulate(ByVal ViewGroup As XtremeCalendarControl.CalendarViewGroup, ByVal Events As XtremeCalendarControl.CalendarEvents)
    Dim pEvent As CalendarEvent
    Dim strData As String
    
    For Each pEvent In Events
        
        pEvent.CustomIcons.RemoveAll
                
       
    Next
End Sub
    
Private Sub CalendarControl_SelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
    If SelType = xtpCalendarSelectionDays Then
        'Debug.Print "SelectionChanged. Day(s)."
    End If
    If SelType = xtpCalendarSelectionEvents Then
        'Debug.Print "SelectionChanged. Event(s)."
        
        Dim HitTest As CalendarHitTestInfo
        Set HitTest = CalendarControl.ActiveView.HitTest
        If Not HitTest.ViewEvent Is Nothing Then
            'ModifyEvent HitTest.ViewEvent.Event, True
        End If
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    CalendarControl.DataProvider.Save
    CalendarControl.DataProvider.Close
End Sub

Public Sub mnuChangeTimeZone()
  '  frmTimeZone.Show vbModal, Me
End Sub

Public Sub mnuDeleteEvent()
    Dim bDeleted As Boolean
    bDeleted = False
    
    If ContextEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence _
        Or ContextEvent.RecurrenceState = xtpCalendarRecurrenceException _
    Then
        frmOccurrenceSeriesChooser.m_bOcurrence = True
        frmOccurrenceSeriesChooser.m_bDeleteRequest = True
        frmOccurrenceSeriesChooser.m_strEventSubject = ContextEvent.Subject
        
        frmOccurrenceSeriesChooser.Show vbModal
        
        If frmOccurrenceSeriesChooser.m_bOK = False Then
            Exit Sub
        ElseIf Not frmOccurrenceSeriesChooser.m_bOcurrence Then
            ' Series
            CalendarControl.DataProvider.DeleteEvent ContextEvent.RecurrencePattern.MasterEvent
            bDeleted = True
        End If
    End If
        
    If Not bDeleted Then
        CalendarControl.DataProvider.DeleteEvent ContextEvent
    End If
    
    CalendarControl.Populate
    CalendarControl.RedrawControl
End Sub

Public Sub mnuNewEvent()
    frmEditEvent.NewEvent
    frmEditEvent.Show vbModal, Me
End Sub

Public Sub mnuOpenEvent()
    ModifyEvent ContextEvent, False
End Sub

'
'
'Public Sub mnOpenColCtas()
'    frmColCtas.Show vbModal, Me
'    frmShortBar.wndShortcutBar.SetFocus
'End Sub



Public Sub mnuTimeScale(newTimeMinutes As Integer)
    CalendarControl.DayView.TimeScale = newTimeMinutes
End Sub

Private Sub ModifyEvent(ModEvent As CalendarEvent, ShowInPane As Boolean)
    If ModEvent.RecurrenceState <> xtpCalendarRecurrenceNotRecurring Then
        
        frmOccurrenceSeriesChooser.m_bOcurrence = True
        frmOccurrenceSeriesChooser.m_bDeleteRequest = False
        frmOccurrenceSeriesChooser.m_strEventSubject = ModEvent.Subject
        
        frmOccurrenceSeriesChooser.Show vbModal
        
        If frmOccurrenceSeriesChooser.m_bOK = False Then
            Exit Sub
        ElseIf frmOccurrenceSeriesChooser.m_bOcurrence Then
            If ModEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
                Set ModEvent = ModEvent.CloneEvent
                ModEvent.MakeAsRException
            End If
        Else
            ' Series
            Set ModEvent = ModEvent.RecurrencePattern.MasterEvent
        End If
    End If

    Set frmEditEvent = New frmEditEvent
    frmEditEvent.ModifyEvent ModEvent
    
    If (ShowInPane) Then
        Dim Pane As Pane
        Set Pane = frmppalN.DockingPaneManager.FindPane(PANE_READING_PANE)
        If Not Pane Is Nothing Then
            frmppalN.DockingPaneManager.DestroyPane Pane
        End If
    Else
        frmEditEvent.BorderStyle = 3
        frmEditEvent.Show vbModal, Me
    End If
End Sub

Function GetMonday(dtDate As Date) As Date
    Dim Day As Long
    Day = WeekDay(dtDate, vbMonday)
    If (Day < 5) Then
        GetMonday = dtDate - Day + 1
    Else
        GetMonday = dtDate - 2
    End If
        
End Function




Public Sub OpenProvider()
  Dim strConnectionString As String
    
    
    Set m_pCustomDataHandler = Nothing
    If CalendarControl.DataProvider.IsOpen Then
        CalendarControl.DataProvider.ClearCache
        CalendarControl.DataProvider.Close
    End If
    
    
    ' MySQL provider
    
    Set m_pCustomDataHandler = New providerMySQL
    '
    strConnectionString = "Provider=Custom;DSN=Ariconta6"
    m_pCustomDataHandler.OpenDB strConnectionString
    m_pCustomDataHandler.SetCalendar CalendarControl

                
    
    'Si pongo PROVIDER=Custom funciona bien, aunque en el connection string le haya dicho la empresa que es
    CalendarControl.SetDataProvider strConnectionString
   ' CalendarControl.SetDataProvider "Provider=Custom;DSN=vAriconta"
    CalendarControl.DataProvider.CacheMode = xtpCalendarDPCacheModeOnRepeat

    CalendarControl.DataProvider.Open

    CalendarControl.Populate
    CalendarControl.EnableReminders True
  ' wndDatePicker.RedrawControl
    
    
    

End Sub


