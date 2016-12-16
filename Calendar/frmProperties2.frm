VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditEvent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva cita"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   Icon            =   "frmProperties2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPrivate 
      Caption         =   "&Privado"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox chkMeeting 
      Caption         =   "Reunión"
      Height          =   195
      Left            =   4200
      TabIndex        =   10
      Top             =   2910
      Width           =   1455
   End
   Begin VB.ComboBox cmbLabel 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ComboBox cmbShowTimeAs 
      Height          =   315
      ItemData        =   "frmProperties2.frx":000C
      Left            =   840
      List            =   "frmProperties2.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2850
      Width           =   1935
   End
   Begin VB.Frame FrameAriadna2 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   4800
      TabIndex        =   27
      Top             =   2100
      Width           =   2535
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedido"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Deshacer enlace ariges"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameAriadna 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   6840
      TabIndex        =   19
      Top             =   -360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton btnCustomProperties 
         Caption         =   "Custom Properties ..."
         Height          =   435
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton btnRecurrence 
         Caption         =   "Recurrence..."
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbSchedule 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "La&bel:"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   885
         Width           =   615
      End
      Begin VB.Label ctrlColor 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5055
         TabIndex        =   24
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblSchedule 
         Caption         =   "Schedule:"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cosas que no vamos a utilizar"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CheckBox chkReminder 
      Caption         =   "Recordar"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmbReminder 
      Height          =   315
      ItemData        =   "frmProperties2.frx":0010
      Left            =   1320
      List            =   "frmProperties2.frx":0012
      TabIndex        =   8
      Text            =   "15 minutes"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtBody 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3240
      Width           =   7095
   End
   Begin VB.CheckBox chkAllDayEvent 
      Caption         =   "Evento de todo el dia"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cmbEndTime 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmbEndDate 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox cmbStartTime 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox cmbStartDate 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   6135
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   240
      X2              =   7320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblShowTimeAs 
      Caption         =   "Mostrar "
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ARIGES"
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   120
      X2              =   7200
      Y1              =   2000
      Y2              =   2000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   7200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblEndTime 
      Caption         =   "Finalización:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1605
      Width           =   855
   End
   Begin VB.Label lblStartTime 
      Caption         =   "Comienzo:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1245
      Width           =   855
   End
   Begin VB.Label lblLocation 
      Caption         =   "Vehículo"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   645
      Width           =   855
   End
   Begin VB.Label lblSubject 
      Caption         =   "Asunto:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   285
      Width           =   855
   End
End
Attribute VB_Name = "frmEditEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'
'
'       .-LABELID lo utilizaremos para saber si la cita enlaza con alguna
'           tabla de NUESTRA gestion
'           0:   NO
'           1:   PEDIDO
'
'       .-ReminderSoundFile lo utilizare para enlazar con los campos de euroges
'         Ejemplo:
'               Si LABELID 1 (pedido) en ReminderSoundFile pondre el campo con el que enlaza

Option Explicit





Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const CB_SETDROPPEDWIDTH = &H160

Dim m_pEditingEvent As CalendarEvent
Dim m_bAddEvent As Boolean


'Enlace con formularios ARIGES
Private vEnlaceAriges As Byte  '0 NO   1: PEdidos
Private vClaveEnlace As String


Private Sub btnCustomProperties_Click()
    If m_pEditingEvent Is Nothing Then
        Exit Sub
    End If
'
'    frmCustomEventProperties.SetEvent m_pEditingEvent
'
'    frmCustomEventProperties.Show vbModal, Me
End Sub

Private Sub btnRecurrence_Click()
'    UpdateEventFromControls
'
'    Set frmEditRecurrence.m_pMasterEvent = m_pEditingEvent.CloneEvent
'    frmEditRecurrence.Show vbModal
'
'    Dim bRecurrenceStateChanged As Boolean
'    bRecurrenceStateChanged = m_pEditingEvent.RecurrenceState <> frmEditRecurrence.m_pMasterEvent.RecurrenceState
'
'    Set m_pEditingEvent = frmEditRecurrence.m_pMasterEvent
'
'    If frmEditRecurrence.m_bUpdateFromEvent Or bRecurrenceStateChanged Then
'        UpdateControlsFromEvent
'    End If

End Sub

Private Sub chkAllDayEvent_Click()

    cmbEndTime.Visible = IIf(chkAllDayEvent.Value = 1, False, True)
    cmbStartTime.Visible = IIf(chkAllDayEvent.Value = 1, False, True)

End Sub

Private Sub chkAllDayEvent_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMeeting_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPrivate_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkReminder_Click()
    cmbReminder.Enabled = IIf(chkReminder.Value > 0, True, False)
    cmbReminder.BackColor = IIf(chkReminder.Value > 0, RGB(255, 255, 255), RGB(210, 210, 210))
End Sub

Private Sub chkReminder_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbEndDate_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbEndTime_Click()
    Dim Index As Long
    Index = InStr(1, cmbEndTime.Text, "(")
    If Index > 0 Then
        cmbEndTime.Text = Left(cmbEndTime.Text, Index - 2)
    End If
    
    
End Sub


Private Sub cmbEndTime_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbLabel_Click()
    Dim pLabel As CalendarEventLabel
    Dim nLabelID As Long
    
    nLabelID = cmbLabel.ItemData(cmbLabel.ListIndex)
    
    Set pLabel = frmMainCalendar.CalendarControl.DataProvider.LabelList.Find(nLabelID)
    If Not pLabel Is Nothing Then
        ctrlColor.BackColor = pLabel.Color
    End If
    
End Sub

Private Sub cmbReminder_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbShowTimeAs_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbStartDate_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbStartTime_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbStartTime_LostFocus()
    UpdateEndTimeCombo
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

Private Sub UpdateEventFromControls()

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
    m_pEditingEvent.Body = txtBody
    m_pEditingEvent.AllDayEvent = chkAllDayEvent.Value = 1
    
    
    '----------------------------------------------------------
    'ENLACE ARIGES
    If m_bAddEvent Then

         m_pEditingEvent.Label = vEnlaceAriges
         m_pEditingEvent.ReminderSoundFile = vClaveEnlace
    Else
        
        If vEnlaceAriges > 0 Then
            'Si vclaveenlace es <0 significa que hemos desenlazado con ariges
            If vClaveEnlace < 0 Then
                m_pEditingEvent.Label = 0
            Else
                m_pEditingEvent.Label = vEnlaceAriges
                m_pEditingEvent.ReminderSoundFile = vClaveEnlace
            End If
        End If
        'm_pEditingEvent.Label = cmbLabel.ItemData(cmbLabel.ListIndex)   'Si enlaza con pedido etc etc
        
    End If
    If m_pEditingEvent.Label = 0 Then m_pEditingEvent.ReminderSoundFile = ""
            
    
    
    'DEpendera de si esta vinculado a ariges (pedido...)
    m_pEditingEvent.BusyStatus = cmbShowTimeAs.ListIndex

    
    
    If cmbSchedule.ListIndex >= 0 And cmbSchedule.ListIndex < cmbSchedule.ListCount Then
        m_pEditingEvent.ScheduleID = cmbSchedule.ItemData(cmbSchedule.ListIndex)
    End If
    
    m_pEditingEvent.PrivateFlag = chkPrivate.Value = 1
    m_pEditingEvent.MeetingFlag = chkMeeting.Value = 1
    
    If Not chkReminder.Value = m_pEditingEvent.Reminder Then m_pEditingEvent.Reminder = chkReminder.Value
    
    
    If chkReminder.Value Then
        If Not Val(cmbReminder.Text) = m_pEditingEvent.ReminderMinutesBeforeStart Then
            m_pEditingEvent.ReminderMinutesBeforeStart = CalcStandardDurations_0m_2wLong(cmbReminder.Text)
        End If
    End If
    
End Sub

Private Sub cmdOk_Click()

    If (Not CheckDates()) Then Exit Sub

    UpdateEventFromControls
    
    If m_bAddEvent Then
        frmMainCalendar.CalendarControl.DataProvider.AddEvent m_pEditingEvent
    Else
        frmMainCalendar.CalendarControl.DataProvider.ChangeEvent m_pEditingEvent
    End If
    
    frmMainCalendar.CalendarControl.Populate

    Unload Me
End Sub

Private Sub UpdateEndTimeCombo()
    On Error GoTo Error
    
    Dim I As Long
    For I = 1 To cmbEndTime.ListCount - 1
        cmbEndTime.RemoveItem 0
    Next I
    
    Dim BeginTime As Date
    BeginTime = TimeValue(cmbStartTime.Text)
    
    cmbEndTime.AddItem BeginTime & " (0 minutes)"
    cmbEndTime.AddItem TimeValue(BeginTime + 1 / 24 / 2) & " (30 minutes)"
    cmbEndTime.AddItem TimeValue(BeginTime + 1 / 24) & " (1 hour)"
    
    For I = 3 To 47
        cmbEndTime.AddItem TimeValue(BeginTime + I / 24 / 2) & " (" & I / 2 & " hours)"
    Next I
    
    Call SendMessage(cmbEndTime.hwnd, CB_SETDROPPEDWIDTH, 200, 0)
    
    
Error:
    
End Sub

Private Sub InitStartTimeCombo()
    On Error GoTo Error
    
    Dim I As Long
    For I = 1 To cmbStartTime.ListCount - 1
        cmbStartTime.RemoveItem 0
    Next I
    
    Dim BeginTime As Date
    BeginTime = #12:00:00 AM#
    
    For I = 1 To 47
        cmbStartTime.AddItem TimeValue(BeginTime + I / 24 / 2)
    Next I
   
Error:
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmPPal.Icon
    InitStartTimeCombo
    
    
    'ASigno las imagenes al toolbar
    With Me.Toolbar1
     '   .ImageList = frmPPal.ImgListPpal
     '   .Buttons(1).Image = 9   'Pedidos
     '   .Buttons(8).Image = 2   'Eliminar  referencia a ARIGES
    End With
    
'    ' Fill Labels Combobox
'    Dim pLabel As CalendarEventLabel
'
'    For Each pLabel In frmMainCalendar.CalendarControl.DataProvider.LabelList
'        cmbLabel.AddItem pLabel.Name
'        cmbLabel.ItemData(cmbLabel.NewIndex) = pLabel.LabelID
'    Next
        
    'ENLACE ARIGES
    cmbLabel.AddItem "Normal"
    cmbLabel.ItemData(cmbLabel.NewIndex) = 0
    cmbLabel.AddItem "Pedido"
    cmbLabel.ItemData(cmbLabel.NewIndex) = 1
        
    cmbLabel.Visible = m_pEditingEvent.Label > 0
    cmbLabel.Enabled = False
    
    ' Fill event Busy Status combobox
    cmbShowTimeAs.AddItem "Libre"
    cmbShowTimeAs.AddItem "Provisional"       'Ponia "Tentative"
    cmbShowTimeAs.AddItem "Ocupado"
    cmbShowTimeAs.AddItem "Fuera de la oficina"
  
    
    
    ' Fill schedules combobox
    Dim pSchedule As CalendarSchedule
    For Each pSchedule In frmMainCalendar.CalendarControl.DataProvider.Schedules
        cmbSchedule.AddItem pSchedule.Name
        cmbSchedule.ItemData(cmbSchedule.NewIndex) = pSchedule.Id
    Next
    
    ' Populate controls with Event properties values
    If Not m_bAddEvent Then
        If m_pEditingEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
            m_pEditingEvent.MakeAsRException
        End If
        
        UpdateControlsFromEvent
    Else
        'Si es nuevo... NO enlaza de salida
        vEnlaceAriges = 0
    End If
    
    ' Fill reminders durations combobox
    FillStandardDurations_0m_2w cmbReminder, False
    
    frmMainCalendar.ModalFormsRunningCounter = frmMainCalendar.ModalFormsRunningCounter + 1

End Sub

Public Sub SetStartEnd(BeginSelection As Date, EndSelection As Date, AllDay As Boolean)
    Dim StartDate As Date, StartTime As Date, EndDate As Date, EndTime As Date

    StartDate = DateValue(BeginSelection)
    StartTime = TimeValue(BeginSelection)

    EndDate = DateValue(EndSelection)
    EndTime = TimeValue(EndSelection)

    If AllDay Then
        cmbEndTime.Visible = False
        cmbStartTime.Visible = False
    
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
    Set m_pEditingEvent = frmMainCalendar.CalendarControl.DataProvider.CreateEvent
    m_bAddEvent = True
    
    Dim BeginSelection As Date, EndSelection As Date, AllDay As Boolean
    frmMainCalendar.CalendarControl.ActiveView.GetSelection BeginSelection, EndSelection, AllDay

    m_pEditingEvent.StartTime = BeginSelection
    m_pEditingEvent.EndTime = EndSelection
    
    SetStartEnd BeginSelection, EndSelection, AllDay
    chkAllDayEvent.Value = IIf(AllDay, 1, 0)
    
    txtSubject = ""

    'cmbShowTimeAs.ListIndex = IIf(AllDay, 0, 2)
    cmbShowTimeAs.ListIndex = 0
    cmbLabel.ListIndex = 0
    If cmbSchedule.ListCount > 0 Then
        Dim nScheduleNr As Integer
        nScheduleNr = frmMainCalendar.CalendarControl.ActiveView.Selection.GroupIndex
        cmbSchedule.ListIndex = nScheduleNr
    End If
    
    chkReminder_Click
    cmbReminder.Text = "15 minutes"
End Sub

Public Sub ModifyEvent(ModEvent As CalendarEvent)
    Set m_pEditingEvent = ModEvent
    m_bAddEvent = False
    
'    txtSubject.Text = m_pEditingEvent.Subject
'    txtBody.Text = m_pEditingEvent.Body
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
    If (m_pEditingEvent.Subject <> "") Then
        Me.Caption = m_pEditingEvent.Subject
    Else
        'Como el texto viene vacio pongo lo que quiero
        Me.Caption = "MODIFICAR CITA"
    End If
    
    
End Sub

Public Sub UpdateControlsFromEvent()
    ' Restore Event base properties
    txtSubject = m_pEditingEvent.Subject
    txtBody = m_pEditingEvent.Body
    txtLocation = m_pEditingEvent.Location
    
    chkAllDayEvent = IIf(m_pEditingEvent.AllDayEvent, 1, 0)
    
    ' Restore Event Label value
    Dim I As Long
    cmbLabel.ListIndex = -1
    For I = 0 To cmbLabel.ListCount - 1
        If cmbLabel.ItemData(I) = m_pEditingEvent.Label Then
            cmbLabel.ListIndex = I
            Exit For
        End If
    Next
    ' Restore Event Schedule value
    For I = 0 To cmbSchedule.ListCount - 1
        If cmbSchedule.ItemData(I) = m_pEditingEvent.ScheduleID Then
            cmbSchedule.ListIndex = I
            Exit For
        End If
    Next
    
    
    cmbShowTimeAs.ListIndex = m_pEditingEvent.BusyStatus
    
    chkPrivate.Value = IIf(m_pEditingEvent.PrivateFlag, 1, 0)
    chkMeeting.Value = IIf(m_pEditingEvent.MeetingFlag, 1, 0)
    
    SetStartEnd m_pEditingEvent.StartTime, m_pEditingEvent.EndTime, m_pEditingEvent.AllDayEvent
    
    If (m_pEditingEvent.Subject <> "") Then
        Me.Caption = " Modificar evento"
    End If

    Dim bDatesVisible As Boolean
    bDatesVisible = m_pEditingEvent.RecurrenceState <> xtpCalendarRecurrenceMaster
       
    lblStartTime.Visible = bDatesVisible
    lblEndTime.Visible = bDatesVisible
    cmbStartDate.Visible = bDatesVisible
    cmbStartTime.Visible = bDatesVisible
    cmbEndDate.Visible = bDatesVisible
    cmbEndTime.Visible = bDatesVisible
    chkAllDayEvent.Visible = bDatesVisible
        
    If bDatesVisible Then
        chkAllDayEvent_Click
    End If
    
    If m_pEditingEvent.Reminder Then
        chkReminder.Value = Checked
        cmbReminder.Text = CalcStandardDurations_0m_2wString(m_pEditingEvent.ReminderMinutesBeforeStart)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMainCalendar.ModalFormsRunningCounter = frmMainCalendar.ModalFormsRunningCounter - 1
End Sub




Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{tab}"
    End If
      
End Sub



Private Sub frmP_DatoSeleccionado2(CadenaSeleccion As String)
    vClaveEnlace = CadenaSeleccion
    vEnlaceAriges = 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
    
        HacerPedido
        
    Case 8
        'Eliminar referencia
        EliminarReferencia
    End Select
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSubject_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub HacerPedido()

    Set frmP = New frmFacEntPedidos
    frmP.DatosADevolverBusqueda2 = "-1"
    If vEnlaceAriges = 0 Then
        'Aun no ha abierto el formulario de pedidos
        If Not m_bAddEvent Then
            If m_pEditingEvent.Label > 0 Then frmP.DatosADevolverBusqueda2 = m_pEditingEvent.ReminderSoundFile
        End If
    Else
        frmP.DatosADevolverBusqueda2 = vClaveEnlace
    End If
    frmP.EsHistorico = False
    frmP.Show vbModal
    Set frmP = Nothing
    
    If vEnlaceAriges > 0 Then Me.cmbLabel.Visible = True
        
  
End Sub

Private Sub EliminarReferencia()
Dim C As String

    C = ""
    If vEnlaceAriges = 0 Then
        'Aun no ha abierto el formulario de pedidos
        If Not m_bAddEvent Then
            If m_pEditingEvent.Label > 0 Then C = m_pEditingEvent.ReminderSoundFile
        End If
    Else
        C = vClaveEnlace
    End If
    If C <> "" Then
        C = "Desea eliminar la referencia de la cita con el pedido: " & C & "?"
        If MsgBox(C, vbQuestion + vbYesNo) = vbYes Then
            If m_bAddEvent Then
                'Si es nueva cita con poner a 0 sobra
                vEnlaceAriges = 0
            Else
                'Si no Tendremos que borrar en la cita
                vEnlaceAriges = 1  'Para que entre a  borrar
                vClaveEnlace = -1
            End If
        End If
    End If
       
End Sub
