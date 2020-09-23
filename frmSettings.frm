VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Search History"
      Height          =   1935
      Left            =   3000
      TabIndex        =   10
      Top             =   240
      Width           =   2295
      Begin VB.PictureBox picCFXPBugFix0 
         BorderStyle     =   0  'None
         Height          =   1685
         Left            =   100
         ScaleHeight     =   1680
         ScaleWidth      =   2100
         TabIndex        =   11
         Top             =   175
         Width           =   2095
         Begin VB.CheckBox ChkSaveHistory 
            Caption         =   "Save"
            Height          =   195
            Left            =   20
            TabIndex        =   14
            Top             =   40
            Width           =   1935
         End
         Begin VB.CommandButton cmdClearHistory 
            Caption         =   "Clear"
            Height          =   315
            Left            =   500
            TabIndex        =   13
            Top             =   1240
            Width           =   1095
         End
         Begin MSComctlLib.Slider SliderHistory 
            Height          =   315
            Left            =   20
            TabIndex        =   12
            Top             =   640
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            LargeChange     =   36
            Min             =   20
            Max             =   200
            SelStart        =   40
            TickFrequency   =   20
            Value           =   40
         End
         Begin VB.Label LblHistory 
            Caption         =   "Size (20-400) :"
            Height          =   255
            Left            =   20
            TabIndex        =   15
            Top             =   400
            Width           =   1935
         End
      End
   End
   Begin VB.CommandButton CmdSetting 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "Apply"
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox ChkSelectWhole 
      Alignment       =   1  'Right Justify
      Caption         =   "Find select whole line"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   510
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Optional Found Grid Headings"
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   2775
      Begin VB.CheckBox ChkGridLine 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Lines"
         Height          =   195
         Left            =   1440
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Routine Name"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Component (If more than one)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Project (If more than one)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CheckBox ChkLaunchStartup 
      Alignment       =   1  'Right Justify
      Caption         =   "Launch On Startup"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox ChkRemFilters 
      Alignment       =   1  'Right Justify
      Caption         =   "Remember Filters"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   315
      Width           =   1935
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OrigShowProj                   As Boolean
Private OrigShowComp                   As Boolean
Private OrigShowRout                   As Boolean
Private OrigLaunchStart                As Boolean
Private OrigRemFilt                    As Boolean
Private OrigSaveHist                   As Boolean
Private OrighistDeep                   As Long
Private origbFindSelectWholeLine       As Boolean
Private ApplyClicked                   As Boolean

Private Sub ChkGridLine_Click()
  If Not bLoadingSettings Then
    bGridlines = ChkGridLine.Value = 1
  End If
End Sub

Private Sub ChkLaunchStartup_Click()
  If Not bLoadingSettings Then
    bLaunchOnStart = ChkLaunchStartup.Value = 1
  End If

End Sub

Private Sub ChkRemFilters_Click()

  If Not bLoadingSettings Then
    bRemFilters = ChkRemFilters.Value = 1
  End If

End Sub

Private Sub ChkSaveHistory_Click()

  If Not bLoadingSettings Then
    bSaveHistory = ChkSaveHistory.Value = 1
  End If

End Sub

Private Sub ChkSelectWhole_Click()

  If Not bLoadingSettings Then
    bFindSelectWholeLine = ChkSelectWhole.Value = 1
  End If

End Sub

Private Sub ChkShow_Click(Index As Integer)

  If Not bLoadingSettings Then
    bShowProject = ChkShow(0).Value = 1
    bShowComponent = ChkShow(1).Value = 1
    bShowRoutine = ChkShow(2).Value = 1
  End If

End Sub

Private Sub cmdClearHistory_Click()

  mobjDoc.ClearHistory

End Sub

Private Sub CmdSetting_Click(Index As Integer)

  ApplyClicked = False
  Select Case Index
   Case 0
    Me.Hide
    mobjDoc.ApplyChanges
   Case 1
    mobjDoc.ApplyChanges
    ApplyClicked = True
   Case 2
    RestoreOriginals
    Me.Hide
  End Select
  SavePropPosition

End Sub

Private Sub Form_Load()

  With Me
    .Left = GetSetting(AppDetails, "Settings", "PropLeft", .Left)
    .Top = GetSetting(AppDetails, "Settings", "PropTop", .Top)
    .Caption = "Properties " & AppDetails
  End With 'Me
  'set safety values for Cancel button
  origbFindSelectWholeLine = bFindSelectWholeLine
  OrigShowProj = bShowProject
  OrigShowComp = bShowComponent
  OrigShowRout = bShowRoutine
  OrigLaunchStart = bLaunchOnStart
  OrigRemFilt = bRemFilters
  OrigSaveHist = bSaveHistory
  OrighistDeep = HistDeep

End Sub

Private Sub Form_Unload(Cancel As Integer)

  If Not ApplyClicked Then
    ' keeps changes if user clicks 'Apply' then uses CaptionBar 'X' button to close
    'otherwise restore
    RestoreOriginals
  End If
  SavePropPosition
  Me.Hide

End Sub

Private Sub RestoreOriginals()

  ChkSelectWhole.Value = Bool2Int(origbFindSelectWholeLine)
  ChkRemFilters.Value = Bool2Int(OrigRemFilt)
  ChkLaunchStartup.Value = Bool2Int(OrigLaunchStart)
  ChkShow(0).Value = Bool2Int(OrigShowProj)
  ChkShow(1).Value = Bool2Int(OrigShowComp)
  ChkShow(2).Value = Bool2Int(OrigShowRout)
  ChkSaveHistory.Value = Bool2Int(OrigSaveHist)
  SliderHistory.Value = OrighistDeep

End Sub

Private Sub SavePropPosition()

  SaveSetting AppDetails, "Settings", "PropLeft", Me.Left
  SaveSetting AppDetails, "Settings", "PropTop", Me.Top

End Sub

Private Sub SliderHistory_Change()

  LblHistory.Caption = "Size (20-400) :" & SliderHistory.Value
  HistDeep = SliderHistory.Value

End Sub

':) Roja's VB Code Fixer V1.1.2 (30/06/2003 3:53:50 PM) 10 + 127 = 137 Lines Thanks Ulli for inspiration and lots of code.
