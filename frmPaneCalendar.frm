VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#17.2#0"; "Codejock.Calendar.v17.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.ShortcutBar.v17.2.0.ocx"
Begin VB.Form frmPaneCalendar2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeCalendarControl.DatePicker DatePicker 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _Version        =   1114114
      _ExtentX        =   4260
      _ExtentY        =   4048
      _StockProps     =   64
      ShowTodayButton =   0   'False
      ShowNoneButton  =   0   'False
      Show3DBorder    =   0
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption MainCaption 
      Height          =   360
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   360
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      GradientColorLight=   15718342
      GradientColorDark=   15718342
      GradientHorizontal=   0   'False
      Expandable      =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ItemCaption 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
      _Version        =   1114114
      _ExtentX        =   5318
      _ExtentY        =   503
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Expandable      =   -1  'True
   End
End
Attribute VB_Name = "frmPaneCalendar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    
    
   
   ' tree.Nodes.Add , , , "Calendario", SHORTCUT_CALENDAR
    
   ' DatePicker.AttachToCalendar frmInbox.CalendarControl
   DatePicker.FirstDayOfWeek = 2
   AsignarCalendar
    
    
        
    DatePicker.AutoSizeRowCol = False
    DatePicker.Width = DatePicker.Width + 30
    
    SetFlatStyle True
    
    UpdateLayout
End Sub

Public Sub AsignarCalendar()
    On Error Resume Next
    
    
    
    If Not frmInbox Is Nothing Then
    
        DatePicker.AttachToCalendar frmInbox.CalendarControl
        DatePicker.FirstDayOfWeek = 2
         frmInbox.CalendarControl.DayView.ShowDay DateTime.Now, True
    End If
    
    Err.Clear
End Sub



Public Sub SetFlatStyle(FlatStyle As Boolean)
      
        Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
        
        MainCaption.GradientColorDark = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
        MainCaption.GradientColorLight = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
        
        ItemCaption.GradientColorDark = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
        ItemCaption.GradientColorLight = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor

End Sub

Private Sub ItemCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ItemCaption.Expanded = Not ItemCaption.Expanded
    UpdateLayout
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ItemCaption.Width = Me.ScaleWidth
    MainCaption.Left = Me.ScaleWidth - MainCaption.Width
    DatePicker.Width = Me.ScaleWidth - MainCaption.Width - 150 - DatePicker.Left
End Sub

Sub UpdateLayout()

    Dim top As Long
    
    top = ItemCaption.top + ItemCaption.Height
  
End Sub

Private Sub MainCaption_ExpandButtonClicked()
    Call frmppalN.ExpandButtonClicked
End Sub
