VERSION 5.00
Begin VB.Form frmEditEvent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Untitled - Event"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbReminder 
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
      ItemData        =   "frmProperties.frx":0000
      Left            =   2160
      List            =   "frmProperties.frx":0002
      TabIndex        =   11
      Text            =   "15 minutes"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CheckBox chkReminder 
      Caption         =   "Recordatorio"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CheckBox chkSoloEstaConta 
      Caption         =   "Solo esta empresa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton btnRecurrence 
      Caption         =   "Recurrence..."
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton btnCustomProperties 
      Caption         =   "Custom Properties ..."
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkMeeting 
      Caption         =   "Reunion"
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
      Left            =   6600
      TabIndex        =   8
      Top             =   1507
      Width           =   1455
   End
   Begin VB.CheckBox chkPrivate 
      Caption         =   "Evento privado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   1500
      Width           =   2415
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   3480
      Width           =   10695
   End
   Begin VB.ComboBox cmbShowTimeAs 
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
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   6240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox chkAllDayEvent 
      Caption         =   "Todo el dia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox cmbEndTime 
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cmbEndDate 
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cmbStartTime 
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
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox cmbStartDate 
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   6360
      Width           =   1215
   End
   Begin VB.ComboBox cmbLabel 
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
      ItemData        =   "frmProperties.frx":0004
      Left            =   7920
      List            =   "frmProperties.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   675
      Width           =   2535
   End
   Begin VB.TextBox txtLocation 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   675
      Width           =   5295
   End
   Begin VB.TextBox txtSubject 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   10680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label ctrlColor 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   10560
      TabIndex        =   24
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblShowTimeAs 
      Caption         =   "Mostrarme como"
      Height          =   195
      Left            =   2400
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   10800
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   10800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblEndTime 
      Caption         =   "Fin"
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
      Left            =   120
      TabIndex        =   19
      Top             =   1965
      Width           =   285
   End
   Begin VB.Label lblStartTime 
      Caption         =   "Comienzo"
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
      Left            =   120
      TabIndex        =   18
      Top             =   1500
      Width           =   930
   End
   Begin VB.Label lblLabel 
      Caption         =   "Etiqueta"
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
      Left            =   6840
      TabIndex        =   15
      Top             =   720
      Width           =   825
   End
   Begin VB.Label lblLocation 
      Caption         =   "Lugar"
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
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   540
   End
   Begin VB.Label lblSubject 
      Caption         =   "Asunto"
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
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "frmEditEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal Wparam As Long, ByVal lParam As Long) As Long
Const CB_SETDROPPEDWIDTH = &H160

Dim m_pEditingEvent As CalendarEvent
Dim m_bAddEvent As Boolean
Public AllDayOverride As Boolean

Private Sub btnCustomProperties_Click()
    If m_pEditingEvent Is Nothing Then
        Exit Sub
    End If
  
    frmCustomEventProperties.SetEvent m_pEditingEvent
  
    frmCustomEventProperties.Show vbModal, Me
End Sub

Private Sub btnRecurrence_Click()
    'Set frmEditRecurrence.m_pMasterEvent = m_pEditingEvent.CloneEvent
    'frmEditRecurrence.Show vbModal
   '
   ' Dim bRecurrenceStateChanged As Boolean
   ' bRecurrenceStateChanged = m_pEditingEvent.RecurrenceState <> frmEditRecurrence.m_pMasterEvent.RecurrenceState
   '
   ' Set m_pEditingEvent = frmEditRecurrence.m_pMasterEvent

    'If frmEditRecurrence.m_bUpdateFromEvent Or bRecurrenceStateChanged Then
    '    UpdateControlsFromEvent
    'End If

End Sub

Private Sub chkAllDayEvent_Click()

    cmbEndTime.visible = IIf(chkAllDayEvent.Value = 1, False, True)
    cmbStartTime.visible = IIf(chkAllDayEvent.Value = 1, False, True)

End Sub

Private Sub chkAllDayEvent_KeyPress(KeyAscii As Integer)

    KEYpressGnral KeyAscii, 3, False


End Sub

Private Sub chkMeeting_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub chkPrivate_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub chkReminder_Click()
    cmbReminder.Enabled = IIf(chkReminder.Value > 0, True, False)
    cmbReminder.BackColor = IIf(chkReminder.Value > 0, RGB(255, 255, 255), RGB(210, 210, 210))
End Sub

Private Sub chkReminder_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub chkSoloEstaConta_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub cmbEndTime_Click()
    Dim Index As Long
    Index = InStr(1, cmbEndTime.Text, "(")
    If Index > 0 Then
        cmbEndTime.Text = Left(cmbEndTime.Text, Index - 2)
    End If
    
    
End Sub


Private Sub cmbLabel_Click()
    Dim pLabel As CalendarEventLabel
    Dim nLabelID As Long
    
    nLabelID = cmbLabel.ItemData(cmbLabel.ListIndex)
    
    Set pLabel = frmInbox.CalendarControl.DataProvider.LabelList.Find(nLabelID)
    If Not pLabel Is Nothing Then
        ctrlColor.BackColor = pLabel.Color
    End If
    
End Sub


Private Sub cmbLabel_KeyPress(KeyAscii As Integer)

    KEYpressGnral KeyAscii, 3, False

End Sub

Private Sub cmbReminder_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub cmbStartDate_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function DateFromString(DatePart As String, TimePart As String) As Date
    Dim dtDatePart As Date, dtTimePart As Date
    dtDatePart = DatePart
    dtTimePart = TimePart
    DateFromString = dtDatePart + dtTimePart
End Function

Function IsDateValid(DatePart As String) As Boolean
    IsDateValid = False
    On Error GoTo Error
    Dim dtDate As Date

    dtDate = DatePart
    IsDateValid = True
Error:
End Function

Private Function CheckDates() As Boolean
    CheckDates = True
    If (Not IsDateValid(cmbStartDate.Text)) Then
        cmbStartDate.SetFocus
        CheckDates = False
        Exit Function
    End If
    If (Not IsDateValid(cmbStartTime.Text)) Then
        cmbStartTime.SetFocus
        CheckDates = False
        Exit Function
    End If
    If (Not IsDateValid(cmbEndDate.Text)) Then
        cmbEndDate.SetFocus
        CheckDates = False
        Exit Function
    End If
    If (Not IsDateValid(cmbEndTime.Text)) Then
        cmbEndTime.SetFocus
        CheckDates = False
        Exit Function
    End If
End Function

Private Sub cmdOk_Click()

    If (Not CheckDates()) Then Exit Sub

    Dim StartTime As Date, EndTime As Date
    StartTime = DateFromString(cmbStartDate.Text, cmbStartTime.Text)
    EndTime = DateFromString(cmbEndDate.Text, cmbEndTime.Text)
    
    If chkAllDayEvent.Value = 1 Then
        If DateDiff("s", TimeValue(EndTime), 0) = 0 Then
            EndTime = EndTime + 1
        End If
    End If
    
    If m_pEditingEvent.RecurrenceState <> xtpCalendarRecurrenceMaster Then
        m_pEditingEvent.StartTime = StartTime
        m_pEditingEvent.EndTime = EndTime
    End If
    
    m_pEditingEvent.Subject = txtSubject.Text
    m_pEditingEvent.Location = txtLocation.Text
    m_pEditingEvent.Body = txtDescription
    m_pEditingEvent.AllDayEvent = chkAllDayEvent.Value = 1
    m_pEditingEvent.Label = cmbLabel.ItemData(cmbLabel.ListIndex)
    m_pEditingEvent.BusyStatus = cmbShowTimeAs.ListIndex
    
    m_pEditingEvent.PrivateFlag = chkPrivate.Value = 1
    m_pEditingEvent.MeetingFlag = chkMeeting.Value = 1
    
     If Not chkReminder.Value = m_pEditingEvent.Reminder Then
        m_pEditingEvent.Reminder = chkReminder.Value
        m_pEditingEvent.ReminderSoundFile = "D:\Backup_10_12\Desktop\mustbuild.wav"
    End If
    
    If chkReminder.Value Then
        If Not Val(cmbReminder.Text) = m_pEditingEvent.ReminderMinutesBeforeStart Then
            m_pEditingEvent.ReminderMinutesBeforeStart = CalcStandardDurations_0m_2wLong(cmbReminder.Text)
        End If
    End If
    
    
    
    If chkSoloEstaConta.Value Then
        m_pEditingEvent.ScheduleID = vEmpresa.codempre
    Else
        m_pEditingEvent.ScheduleID = 0
    End If
    
    If m_bAddEvent Then
        frmInbox.CalendarControl.DataProvider.AddEvent m_pEditingEvent
    Else
        frmInbox.CalendarControl.DataProvider.ChangeEvent m_pEditingEvent
    End If
    
    frmInbox.CalendarControl.Populate

    Unload Me
End Sub

Private Sub UpdateEndTimeCombo()
    On Error GoTo Error
    
    Dim i As Long
    For i = 1 To cmbEndTime.ListCount - 1
        cmbEndTime.RemoveItem 0
    Next i
    
    Dim BeginTime As Date
    BeginTime = TimeValue(cmbStartTime.Text)
    
    cmbEndTime.AddItem BeginTime & " (0 minutes)"
    cmbEndTime.AddItem TimeValue(BeginTime + 1 / 24 / 2) & " (30 minutes)"
    cmbEndTime.AddItem TimeValue(BeginTime + 1 / 24) & " (1 hour)"
    
    For i = 3 To 47
        cmbEndTime.AddItem TimeValue(BeginTime + i / 24 / 2) & " (" & i / 2 & " hours)"
    Next i
    
    Call SendMessage(cmbEndTime.hwnd, CB_SETDROPPEDWIDTH, 200, 0)
    
    
Error:
    
End Sub

Private Sub Form_Load()
Dim J As Integer
Dim C As String

    'Me.Icon = frmPpal.Icon  FALTA###
    '===============================
    Dim pLabel As CalendarEventLabel
    

    J = 0
    For Each pLabel In frmInbox.CalendarControl.DataProvider.LabelList
        J = J + 1
        C = RecuperaValor(TextosLabelEspanol, CInt(J))
        If C = "" Then C = pLabel.Name
        cmbLabel.AddItem C
        cmbLabel.ItemData(cmbLabel.NewIndex) = pLabel.LabelID
    Next
        
    cmbShowTimeAs.AddItem "Free"
    cmbShowTimeAs.AddItem "Tentative"
    cmbShowTimeAs.AddItem "Busy"
    cmbShowTimeAs.AddItem "Out of Office"
    
    If Not m_bAddEvent Then
        If Not m_pEditingEvent Is Nothing Then
            If m_pEditingEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
                m_pEditingEvent.MakeAsRException
            End If
        
            UpdateControlsFromEvent
        End If
    Else
        Caption = "Nuevo"
    End If
        
    ' Fill reminders durations combobox
    
    FillStandardDurations_0m_2w cmbReminder, False
    If m_bAddEvent Then cmbReminder.Text = "15 minutos"
End Sub




Public Sub SetStartEnd(BeginSelection As Date, EndSelection As Date, AllDay As Boolean)
    Dim StartDate As Date, StartTime As Date, EndDate As Date, EndTime As Date

    StartDate = DateValue(BeginSelection)
    StartTime = TimeValue(BeginSelection)

    EndDate = DateValue(EndSelection)
    EndTime = TimeValue(EndSelection)

    If AllDay Then
        cmbEndTime.visible = False
        cmbStartTime.visible = False
    
        If DateDiff("s", EndTime, 0) = 0 Then
            EndDate = EndDate - 1
        End If
    End If
    
    cmbStartDate.Text = StartDate
    cmbStartTime.Text = StartTime
    
    UpdateEndTimeCombo

    cmbEndDate.Text = EndDate
    cmbEndTime.Text = EndTime
    
 
End Sub


Public Sub NewEvent()
    Set m_pEditingEvent = frmInbox.CalendarControl.DataProvider.CreateEvent
    m_bAddEvent = True
    
    Dim BeginSelection As Date, EndSelection As Date, AllDay As Boolean
    frmInbox.CalendarControl.ActiveView.getSelection BeginSelection, EndSelection, AllDay

    SetStartEnd BeginSelection, EndSelection, AllDay
    
    If AllDayOverride Then
        AllDay = True
    End If
    
    chkAllDayEvent.Value = IIf(AllDay, 1, 0)
    
    txtSubject = "Nuevo evento"

    cmbShowTimeAs.ListIndex = IIf(AllDay, 0, 2)
    cmbLabel.ListIndex = 0
    
    AllDayOverride = False
    
End Sub

Public Sub ModifyEvent(ModEvent As CalendarEvent)
    Set m_pEditingEvent = ModEvent
    m_bAddEvent = False
    
'    txtSubject.Text = m_pEditingEvent.Subject
'    txtDescription.Text = m_pEditingEvent.Description
'    txtLocation.Text = m_pEditingEvent.Location
'
'    chkAllDayEvent.Value = IIf(m_pEditingEvent.AllDayEvent, 1, 0)
'
'    Dim i As Long
'    For i = 0 To cmbLabel.ListCount - 1
'        If cmbLabel.ItemData(i) = m_pEditingEvent.Label Then
'            cmbLabel.ListIndex = i
'            Exit For
'        End If
'    Next
'
'    cmbShowTimeAs.ListIndex = m_pEditingEvent.BusyStatus
'
'    chkPrivate.Value = IIf(m_pEditingEvent.PrivateFlag, 1, 0)
'    chkMeeting.Value = IIf(m_pEditingEvent.MeetingFlag, 1, 0)
'
'    SetStartEnd m_pEditingEvent.StartTime, m_pEditingEvent.EndTime, m_pEditingEvent.AllDayEvent
'
    If (m_pEditingEvent.Subject <> "") Then Me.Caption = m_pEditingEvent.Subject & " - Evento"
    
End Sub

Public Sub UpdateControlsFromEvent()

    txtSubject = m_pEditingEvent.Subject
    txtDescription = m_pEditingEvent.Body
    txtLocation = m_pEditingEvent.Location
    
    chkAllDayEvent = IIf(m_pEditingEvent.AllDayEvent, 1, 0)
    
    Dim i As Long
    For i = 0 To cmbLabel.ListCount - 1
        If cmbLabel.ItemData(i) = m_pEditingEvent.Label Then
            cmbLabel.ListIndex = i
            Exit For
        End If
    Next
    
    cmbShowTimeAs.ListIndex = m_pEditingEvent.BusyStatus
    
    chkPrivate.Value = IIf(m_pEditingEvent.PrivateFlag, 1, 0)
    chkMeeting.Value = IIf(m_pEditingEvent.MeetingFlag, 1, 0)
    Me.chkSoloEstaConta.Value = IIf(m_pEditingEvent.ScheduleID = 0, 0, 1)
    
    
    
    SetStartEnd m_pEditingEvent.StartTime, m_pEditingEvent.EndTime, m_pEditingEvent.AllDayEvent
    
    
    
    chkReminder.Value = IIf(m_pEditingEvent.Reminder, 1, 0)
    
    
    If chkReminder.Value Then
        If Not Val(cmbReminder.Text) = m_pEditingEvent.ReminderMinutesBeforeStart Then
            cmbReminder.Text = CalcStandardDurations_0m_2wString(m_pEditingEvent.ReminderMinutesBeforeStart)
        End If
    End If
    
    
    
    
    
    
    
    
    
    
    If (m_pEditingEvent.Subject <> "") Then
        Me.Caption = m_pEditingEvent.Subject & " - Evento"
    End If

    Dim bDatesVisible As Boolean
    bDatesVisible = m_pEditingEvent.RecurrenceState <> xtpCalendarRecurrenceMaster
       
    lblStartTime.visible = bDatesVisible
    lblEndTime.visible = bDatesVisible
    cmbStartDate.visible = bDatesVisible
    cmbStartTime.visible = bDatesVisible
    cmbEndDate.visible = bDatesVisible
    cmbEndTime.visible = bDatesVisible
    chkAllDayEvent.visible = bDatesVisible
        
    If bDatesVisible Then
        chkAllDayEvent_Click
    End If
    
End Sub


Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub txtSubject_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub
