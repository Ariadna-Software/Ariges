VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReminders 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminders"
   ClientHeight    =   4305
   ClientLeft      =   3045
   ClientTop       =   3330
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   9
      Top             =   4720
      Width           =   1215
   End
   Begin VB.ComboBox cmbSnooze 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmReminders.frx":0000
      Left            =   120
      List            =   "frmReminders.frx":0002
      TabIndex        =   6
      Top             =   3840
      Width           =   3735
   End
   Begin VB.CommandButton btnSnooze 
      Caption         =   "&Repetir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5760
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton btnDismiss 
      Caption         =   "&Descartar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnOpenItem 
      Caption         =   "&Abrir evento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton btnDismissAll 
      Caption         =   "Descartar &todo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin MSComctlLib.ListView ctrlReminders 
      Height          =   1785
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   3149
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
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
         Text            =   "Evento"
         Object.Width           =   7144
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vencido"
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.Label txtDescription2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label Label2 
      Caption         =   "Repetir de nuevo en:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label txtDescription1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmReminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CuantosAvisos As Integer


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
    
    CuantosAvisos = ctrlReminders.ListItems.Count
    If ctrlReminders.ListItems.Count = 0 Then
        Unload Me
    End If
End Sub

Private Sub UpdateFromManager()
    ctrlReminders.ListItems.Clear
        
    Dim pRemI As CalendarReminder
    Dim pEventI As CalendarEvent
    Dim pItemI
        
    For Each pRemI In frmInbox.CalendarControl.Reminders
        Set pEventI = pRemI.Event
        Set pItemI = ctrlReminders.ListItems.Add()
        
        pItemI.Text = pEventI.Subject
             
        Dim nMinutes As Long, strDueIn As String
        nMinutes = DateDiff("n", Now, pEventI.StartTime)
        
        If nMinutes > 0 Then
            strDueIn = FormatTimeDuration(nMinutes, True)
        Else
            strDueIn = FormatTimeDuration(-1 * nMinutes, True) & " atrasado"
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
            txtDescription2.Caption = "0 eventos seleccionados"
        Else
            txtDescription2.Caption = "No hay eventos para mostrar."
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
        Set pRem = frmInbox.CalendarControl.Reminders(ctrlReminders.SelectedItem.Index - 1)
        
        txtDescription1.Caption = pRem.Event.Subject
        txtDescription2.Caption = "Comienzo:  " & FormatDateTime(pRem.Event.StartTime)
        
        If (pRem.MinutesBeforeStart < 5) Then
            cmbSnooze.Text = "5 minutos"
        Else
            cmbSnooze.Text = FormatTimeDuration(pRem.MinutesBeforeStart, False)
        End If
    End If
    
    Caption = ctrlReminders.ListItems.Count & " Alerta" & IIf(ctrlReminders.ListItems.Count > 1, "s", "")
End Sub

Private Sub btnDismiss_Click()
    If ctrlReminders.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim pRem As CalendarReminder
    Dim nIndex As Long
    nIndex = ctrlReminders.SelectedItem.Index
    Set pRem = frmInbox.CalendarControl.Reminders(nIndex - 1)
    pRem.Dismiss
End Sub

Private Sub btnDismissAll_Click()
    frmInbox.CalendarControl.Reminders.DismissAll
End Sub

Private Sub btnOpenItem_Click()
    If ctrlReminders.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim pRem As CalendarReminder
    Dim nIndex As Long
    nIndex = ctrlReminders.SelectedItem.Index
    Set pRem = frmInbox.CalendarControl.Reminders(nIndex - 1)
    
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
    Set pRem = frmInbox.CalendarControl.Reminders(nIndex - 1)
    pRem.Snooze nMinutes
End Sub




Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub ctrlReminders_ItemClick(ByVal Item As MSComctlLib.ListItem)
UpdateControlsBySelection
End Sub

Private Sub Form_Activate()
    If ctrlReminders.ListItems.Count = 0 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmppalN.Icon
    FillStandardDurations_0m_2w cmbSnooze, True
End Sub

