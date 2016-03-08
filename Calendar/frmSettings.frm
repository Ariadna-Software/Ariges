VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Day/Week/Month View"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   3000
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
      Begin VB.ComboBox cmbWeeksCount 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblWeeksCount 
         Caption         =   "Contar semana"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Skin Office 2007"
      Height          =   735
      Left            =   3240
      TabIndex        =   35
      Top             =   3960
      Width           =   2775
      Begin VB.CheckBox chkSkinOffice2007 
         Caption         =   "Aplicar Skin Off. 2007"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame frmAdditionalOpt 
      Caption         =   "Opciones adicionales"
      Height          =   2415
      Left            =   0
      TabIndex        =   27
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox chkMVShowEndTimeAlways 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar siempre hora fin"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkMVShowStartTimeAlways 
         Caption         =   "Mostrar siempre hora inicio"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CheckBox chkWVShowEndTimeAlways 
         Caption         =   "Mostrar siempre hora fin"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkWVShowStartTimeAlways 
         Caption         =   "Mostrar siempre hora inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Mes vista:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Semana vista:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Otros"
      Height          =   2055
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   2955
      Begin VB.CheckBox chkComperssWeekendDays 
         Caption         =   "Comprimir dias fin de semana"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chkEnableReminders 
         Caption         =   "Habilitar avisos"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cmbToolTipsMode 
         Height          =   315
         ItemData        =   "frmSettings.frx":0000
         Left            =   1320
         List            =   "frmSettings.frx":0002
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ToolTips:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame frmMonth 
      Caption         =   "Mes vista"
      Height          =   975
      Left            =   3240
      TabIndex        =   21
      Top             =   2880
      Width           =   2775
      Begin VB.CheckBox chkShowEndTimeMonth 
         Caption         =   "Mostrar hora final"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkShowTimeAsClockMonth 
         Caption         =   "Mostrar hora como un reloj"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frmWeek 
      Caption         =   "Semana vista"
      Height          =   975
      Left            =   3240
      TabIndex        =   20
      Top             =   1800
      Width           =   2775
      Begin VB.CheckBox chkShowTimeAsClockWeek 
         Caption         =   "Mostrar hora como un reloj"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox chkShowEndTimeWeek 
         Caption         =   "Mostrar hora fin"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   660
         Width           =   1935
      End
   End
   Begin VB.Frame frmWorkWeek 
      Caption         =   "Semana laboral"
      Height          =   1575
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6015
      Begin VB.ComboBox cmbEndTime 
         Height          =   315
         Left            =   4440
         TabIndex        =   16
         Text            =   "cmbEndTime"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbStartTime 
         Height          =   315
         ItemData        =   "frmSettings.frx":0004
         Left            =   4440
         List            =   "frmSettings.frx":0006
         TabIndex        =   15
         Text            =   "cmbStartTime"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbFirstDayOfWeek 
         Height          =   315
         ItemData        =   "frmSettings.frx":0008
         Left            =   1560
         List            =   "frmSettings.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Sab"
         Height          =   195
         Index           =   6
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Vie"
         Height          =   195
         Index           =   5
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Jue"
         Height          =   195
         Index           =   4
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Mier"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Mar"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Lun"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Dom"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblEndTime 
         Caption         =   "Hora fin:"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblStartTime 
         Caption         =   "Hora inicio:"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label lblFirstDayOfWeek 
         Caption         =   "Primer dia semana"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   765
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame frmDay 
      Caption         =   "Dia vista"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   3135
      Begin VB.CommandButton cmdTimeZone 
         Caption         =   "Zona horaria.."
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cmbTimeScale 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTimeScale 
         AutoSize        =   -1  'True
         Caption         =   "Escala de tiempo:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Property Get CalendarControl() As CalendarControl
    Set CalendarControl = frmMainCalendar.CalendarControl
End Property

Private Sub chkMVShowStartTimeAlways_Click()
    chkMVShowEndTimeAlways.Enabled = chkMVShowStartTimeAlways.Value <> 0
    If Not chkMVShowEndTimeAlways.Enabled Then
        chkMVShowEndTimeAlways.Value = 0
    End If
    
End Sub

Private Sub chkWVShowStartTimeAlways_Click()
     chkWVShowEndTimeAlways.Enabled = chkWVShowStartTimeAlways.Value <> 0
     If Not chkWVShowEndTimeAlways.Enabled Then
        chkWVShowEndTimeAlways.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    'Guardo cambios
    chkWVShowEndTimeAlways.Value = chkShowEndTimeWeek.Value
    chkMVShowEndTimeAlways.Value = chkShowEndTimeMonth.Value
    
    ApplySettings
    CalendarControl.Populate
    
    
    Unload Me
End Sub

Sub AddTimeScale(TimeScale As Long)
    cmbTimeScale.AddItem TimeScale & " minutes"
    cmbTimeScale.ItemData(cmbTimeScale.ListCount - 1) = TimeScale

    If CalendarControl.DayView.TimeScale = TimeScale Then cmbTimeScale.ListIndex = cmbTimeScale.ListCount - 1
End Sub

Sub AddCalendarDay(Index As Long, Day As CalendarWeekDay, Caption As String, FirstDayOfTheWeek As Long)
    chkWorkDay(Index).Value = IIf(CalendarControl.Options.WorkWeekMask And Day, 1, 0)
    
    cmbFirstDayOfWeek.AddItem Caption
    If (CalendarControl.Options.FirstDayOfTheWeek = FirstDayOfTheWeek) Then cmbFirstDayOfWeek.ListIndex = Index
    
End Sub

Private Sub cmdTimeZone_Click()
'    If frmMainCalendar.g_bUseBuiltInCalendarDialogs Then
'        Dim dlgCalendar As New CalendarDialogs
'        dlgCalendar.ParentHWND = Me.hwnd
'        dlgCalendar.Calendar = CalendarControl
'
'        dlgCalendar.ShowTimeScaleProperties
'        Exit Sub
'    End If
'
'    frmTimeZone.Show vbModal, Me
End Sub

Private Sub Form_Load()
    AddTimeScale 5
    AddTimeScale 6
    AddTimeScale 10
    AddTimeScale 15
    AddTimeScale 30
    AddTimeScale 60
    
    Dim WorkWeekMask As CalendarWeekDay
    WorkWeekMask = CalendarControl.Options.WorkWeekMask
    
    AddCalendarDay 0, xtpCalendarDaySunday, "Domingo", 1
    AddCalendarDay 1, xtpCalendarDayMonday, "Lunes", 2
    AddCalendarDay 2, xtpCalendarDayTuesday, "Martes", 3
    AddCalendarDay 3, xtpCalendarDayWednesday, "Miércoles", 4
    AddCalendarDay 4, xtpCalendarDayThursday, "Jueves", 5
    AddCalendarDay 5, xtpCalendarDayFriday, "Viernes", 6
    AddCalendarDay 6, xtpCalendarDaySaturday, "Sábado", 7
    
    cmbStartTime.Text = CalendarControl.Options.WorkDayStartTime
    cmbEndTime.Text = CalendarControl.Options.WorkDayEndTime
    
    chkShowTimeAsClockWeek.Value = IIf(CalendarControl.Options.WeekViewShowTimeAsClocks, 1, 0)
    chkShowEndTimeWeek.Value = IIf(CalendarControl.Options.WeekViewShowEndDate, 1, 0)
    
    chkShowTimeAsClockMonth.Value = IIf(CalendarControl.Options.MonthViewShowTimeAsClocks, 1, 0)
    chkShowEndTimeMonth.Value = IIf(CalendarControl.Options.MonthViewShowEndDate, 1, 0)
    chkComperssWeekendDays.Value = IIf(CalendarControl.Options.MonthViewCompressWeekendDays, 1, 0)
    
    cmbWeeksCount.AddItem "2"
    cmbWeeksCount.AddItem "3"
    cmbWeeksCount.AddItem "4"
    cmbWeeksCount.AddItem "5"
    cmbWeeksCount.AddItem "6"
    cmbWeeksCount.ListIndex = CalendarControl.MonthView.WeeksCount - 2
    
    cmbToolTipsMode.AddItem "Standard"
    cmbToolTipsMode.ItemData(cmbToolTipsMode.NewIndex) = 0
    
    cmbToolTipsMode.AddItem "Personalizado"
    cmbToolTipsMode.ItemData(cmbToolTipsMode.NewIndex) = 1
    
    cmbToolTipsMode.AddItem "Deshabilitado"
    cmbToolTipsMode.ItemData(cmbToolTipsMode.NewIndex) = 2
    
    Dim I As Long
    For I = 0 To 2
        If cmbToolTipsMode.ItemData(I) = frmMainCalendar.ToolTips_Mode Then
            cmbToolTipsMode.ListIndex = I
            Exit For
        End If
    Next
    
    '---------------------------------------------------------------
    chkWVShowStartTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptWeekViewShowStartTimeAlways), 1, 0)
    chkWVShowEndTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptWeekViewShowEndTimeAlways), 1, 0)
    
    chkMVShowStartTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptMonthViewShowStartTimeAlways), 1, 0)
    chkMVShowEndTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptMonthViewShowEndTimeAlways), 1, 0)
                                    
                                    
    chkEnableReminders.Value = BooleanToBin(CalendarControl.IsRemindersEnabled)
    chkMVShowStartTimeAlways.Value = 1
    chkWVShowStartTimeAlways.Value = 1
    chkMVShowStartTimeAlways_Click
    chkWVShowStartTimeAlways_Click
        
    Me.chkSkinOffice2007.Value = BooleanToBin(frmMainCalendar.Skin2007)
        
    'En el cmdOK esta al reves
    chkShowEndTimeWeek.Value = chkWVShowEndTimeAlways.Value
    chkShowEndTimeMonth.Value = chkMVShowEndTimeAlways.Value
        
        
    frmMainCalendar.ModalFormsRunningCounter = frmMainCalendar.ModalFormsRunningCounter + 1
End Sub


Sub ApplyCalendarDay(Index As Long, Day As CalendarWeekDay, FirstDayOfTheWeek As Long)
    
    If (chkWorkDay(Index).Value) Then CalendarControl.Options.WorkWeekMask = CalendarControl.Options.WorkWeekMask Or Day
    
    If (cmbFirstDayOfWeek.ListIndex = Index) Then CalendarControl.Options.FirstDayOfTheWeek = FirstDayOfTheWeek
    
End Sub


Sub ApplySettings()
    Dim eViewType As Long
    eViewType = CalendarControl.ViewType
    
    CalendarControl.Options.WorkWeekMask = 0

    ApplyCalendarDay 0, xtpCalendarDaySunday, 1
    ApplyCalendarDay 1, xtpCalendarDayMonday, 2
    ApplyCalendarDay 2, xtpCalendarDayTuesday, 3
    ApplyCalendarDay 3, xtpCalendarDayWednesday, 4
    ApplyCalendarDay 4, xtpCalendarDayThursday, 5
    ApplyCalendarDay 5, xtpCalendarDayFriday, 6
    ApplyCalendarDay 6, xtpCalendarDaySaturday, 7
    
    
    CalendarControl.DayView.TimeScale = cmbTimeScale.ItemData(cmbTimeScale.ListIndex)
    
    CalendarControl.Options.WeekViewShowTimeAsClocks = chkShowTimeAsClockWeek.Value
    CalendarControl.Options.WeekViewShowEndDate = chkShowEndTimeWeek.Value
    
    CalendarControl.Options.MonthViewShowTimeAsClocks = chkShowTimeAsClockMonth.Value
    CalendarControl.Options.MonthViewShowEndDate = chkShowEndTimeMonth.Value
    CalendarControl.Options.MonthViewCompressWeekendDays = chkComperssWeekendDays.Value
    
    CalendarControl.MonthView.WeeksCount = cmbWeeksCount.ListIndex + 2
    
    CalendarControl.Options.WorkDayStartTime = TimeValue(cmbStartTime.Text)
    CalendarControl.Options.WorkDayEndTime = TimeValue(cmbEndTime.Text)
    
    
    
    '---------------------------------------------------------------
    CalendarControl.Options.AdditionalOptionsFlags.Flags = 0
    
    If chkWVShowStartTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptWeekViewShowStartTimeAlways
    Else
        
    End If
    
    If chkWVShowEndTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptWeekViewShowEndTimeAlways
    End If
    
    If chkMVShowStartTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptMonthViewShowStartTimeAlways
    End If
        
    If chkMVShowEndTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptMonthViewShowEndTimeAlways
    End If


    CalendarControl.EnableReminders BinToBoolean(chkEnableReminders.Value)
    
    frmMainCalendar.Skin2007 = BinToBoolean(chkSkinOffice2007.Value)
    
    'to apply WorkWeekMask changes
    CalendarControl.ViewType = eViewType
        
    frmMainCalendar.ToolTips_Mode = cmbToolTipsMode.ItemData(cmbToolTipsMode.ListIndex)
    frmMainCalendar.CambioConfiguracion = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMainCalendar.ModalFormsRunningCounter = frmMainCalendar.ModalFormsRunningCounter - 1
End Sub


