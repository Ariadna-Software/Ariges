VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReminders 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avisos pendientes"
   ClientHeight    =   4305
   ClientLeft      =   3045
   ClientTop       =   3330
   ClientWidth     =   6480
   Icon            =   "frmReminders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ListView ctrlReminders 
      Height          =   1815
      Left            =   600
      TabIndex        =   8
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Asunto"
         Object.Width           =   5645
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vence en"
         Object.Width           =   3775
      EndProperty
   End
   Begin VB.ComboBox cmbSnooze 
      Height          =   315
      ItemData        =   "frmReminders.frx":000C
      Left            =   120
      List            =   "frmReminders.frx":000E
      TabIndex        =   6
      Top             =   3840
      Width           =   4935
   End
   Begin VB.CommandButton btnSnooze 
      Caption         =   "&Recordar"
      Default         =   -1  'True
      Height          =   315
      Left            =   5160
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton btnDismiss 
      Caption         =   "&Descartar"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnOpenItem 
      Caption         =   "&Abrir evento"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnDismissAll 
      Caption         =   "Descartar &todos"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmReminders.frx":0010
      Top             =   240
      Width           =   480
   End
   Begin VB.Label txtDescription2 
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Recordármelo más tarde"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   4815
   End
   Begin VB.Label txtDescription1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmReminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub OnReminders(ByVal Action As XtremeCalendarControl.CalendarRemindersAction, ByVal Reminder As XtremeCalendarControl.CalendarReminder)
    If Action = xtpCalendarRemindersFire Or Action = xtpCalendarReminderSnoozed Or _
       Action = xtpCalendarReminderDismissed Or Action = xtpCalendarReminderDismissedAll _
    Then
        UpdateFromManager
        UpdateControlsBySelection
        
    ElseIf Action = xtpCalendarRemindersMonitoringStopped Then
        ctrlReminders.ListItems.Clear
        UpdateControlsBySelection
    End If
    
    If ctrlReminders.ListItems.Count = 0 Then
        Unload Me
    Else
        Sonido
    End If
End Sub

Private Sub Sonido()
    On Error Resume Next
    Beep
    Beep
    Err.Clear
End Sub

Private Sub UpdateFromManager()
    ctrlReminders.ListItems.Clear
        
    Dim pRemI As CalendarReminder
    Dim pEventI As CalendarEvent
    Dim pItemI As ListItem
        
    For Each pRemI In frmMainCalendar.CalendarControl.Reminders
        Set pEventI = pRemI.Event
        Set pItemI = ctrlReminders.ListItems.Add("1")
        
        pItemI.Text = pEventI.Subject
             
        Dim nMinutes As Long, strDueIn As String
        nMinutes = DateDiff("n", Now, pEventI.StartTime)
        
        If nMinutes > 0 Then
            strDueIn = FormatTimeDuration(nMinutes, True)
        Else
            strDueIn = FormatTimeDuration(-1 * nMinutes, True) & " que ha vencido"
        End If
        
        pItemI.SubItems(1) = strDueIn
    Next
    
End Sub

Private Sub UpdateControlsBySelection()
    Dim bEnabled As Boolean
    bEnabled = False
    
    If ctrlReminders.SelectedItem Is Nothing Then
        txtDescription1.Caption = ""
        If ctrlReminders.ListItems.Count > 0 Then
            txtDescription2.Caption = "No hay avisos seleccionados"
        Else
            txtDescription2.Caption = "No hay avisos para mostrar."
        End If
    Else
        bEnabled = True
    End If
    
    btnDismissAll.Enabled = bEnabled
    btnDismiss.Enabled = bEnabled
    btnOpenItem.Enabled = bEnabled
    btnSnooze.Enabled = bEnabled
    cmbSnooze.Enabled = bEnabled
    
    Dim pRem As CalendarReminder
        
    If bEnabled Then
        Set pRem = frmMainCalendar.CalendarControl.Reminders(ctrlReminders.SelectedItem.Index - 1)
        
        txtDescription1.Caption = pRem.Event.Subject
        txtDescription2.Caption = "Hora inicio:  " & FormatDateTime(pRem.Event.StartTime)
        
        If (pRem.MinutesBeforeStart < 5) Then
            cmbSnooze.Text = "5 minutos"
        Else
            cmbSnooze.Text = FormatTimeDuration(pRem.MinutesBeforeStart, False)
        End If
    End If
    
    Caption = ctrlReminders.ListItems.Count & " cita" & IIf(ctrlReminders.ListItems.Count > 1, "s", "")
End Sub

Private Sub btnDismiss_Click()
    If ctrlReminders.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim pRem As CalendarReminder
    Dim nIndex As Long
    nIndex = ctrlReminders.SelectedItem.Index
    Set pRem = frmMainCalendar.CalendarControl.Reminders(nIndex - 1)
    pRem.Dismiss
End Sub

Private Sub btnDismissAll_Click()
    frmMainCalendar.CalendarControl.Reminders.DismissAll
End Sub

Private Sub btnOpenItem_Click()
    If ctrlReminders.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim pRem As CalendarReminder
    Dim nIndex As Long
    nIndex = ctrlReminders.SelectedItem.Index
    Set pRem = frmMainCalendar.CalendarControl.Reminders(nIndex - 1)
    
    Dim frmProperties As New frmEditEvent
    frmProperties.ModifyEvent pRem.Event
    frmProperties.Show vbModal, Me
End Sub

Private Sub btnSnooze_Click()
    If ctrlReminders.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim nMinutes As Long
    ParseTimeDuration cmbSnooze.Text, nMinutes

    Dim pRem As CalendarReminder
    Dim nIndex As Long
    nIndex = ctrlReminders.SelectedItem.Index
    Set pRem = frmMainCalendar.CalendarControl.Reminders(nIndex - 1)
    pRem.Snooze nMinutes
End Sub





Private Sub ctrlReminders_ItemClick(ByVal Item As MSComctlLib.ListItem)
UpdateControlsBySelection
End Sub

Private Sub Form_Load()
    FillStandardDurations_0m_2w cmbSnooze, True
    Me.Icon = frmPpal.Icon
End Sub
