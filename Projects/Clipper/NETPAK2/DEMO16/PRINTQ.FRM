VERSION 4.00
Begin VB.Form frmPrintQ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Capture LPT Port"
   ClientHeight    =   3735
   ClientLeft      =   3480
   ClientTop       =   2100
   ClientWidth     =   2190
   Height          =   4140
   Icon            =   "PRINTQ.frx":0000
   Left            =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Top             =   1755
   Width           =   2310
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdDetach 
      Caption         =   "Dettach Print Queue"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Attach Print Queue"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox cmbLPT 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox cmbPrintQueues 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "LPT Port"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Print &Queue"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrintQ"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim ccode As Integer

Const RELEASE_JOB = 4
Const SUPPRESS_FORM_FEED = 8
Const NOTIFY = 16
Const INTERPRET_CODES = 64
Const PRINT_BANNER = 128


Private Sub cmdAttach_Click()
  Dim iLPT As Integer, SetFlags As SetCaptureFlags
  Dim flags As Integer
  flags = 0
  ' Assign LPT port to print queue
  '--------------------------
  iLPT = Val(cmbLPT) - 1
  If (iLPT = -1) Or (Len(cmbprintqueues) = 0) Then
    MsgBox "Data not correct.", 16, "No capture"
  Else
    ccode% = NW_EndSPLPTCap(iLPT)
    ccode% = NW_SetCapturePrintQ(iLPT, cmbprintqueues)
    
    SetFlags.status = Chr$(0)
    flags = flags + RELEASE_JOB
    flags = flags + SUPPRESS_FORM_FEED
    flags = flags + NOTIFY
    flags = flags + INTERPRET_CODES
    flags = flags + PRINT_BANNER
    SetFlags.flags = Chr$(flags)
    SetFlags.tab_size = Chr$(0)
    SetFlags.number_of_copies = Chr$(1)
    SetFlags.banner_text = ""
    SetFlags.local_lpt_device = iLPT
    SetFlags.flush_capture_timeout_count = 30
    
    ccode% = NW_SetDefCapFlags%(SetFlags)
    ccode% = NW_StartSPLPTCap(iLPT)
    
    If ccode% = 0 Then
      MsgBox "Port Captured.", 64, "Done"
    End If
  End If
End Sub

Private Sub cmdDetach_Click()
  Dim iLPT As Integer
  iLPT = Val(cmbLPT) - 1
  If iLPT > -1 Then
    ccode% = NW_EndSPLPTCap(iLPT)
    If ccode% = 0 Then
      MsgBox "Detachment Successful.", 64, "Done"
    Else
      MsgBox "Can't Detach LPT Port.", 16, "Network Error"
    End If
  Else
    MsgBox "Data not correct.", 16, "Detach not performed"
  End If
End Sub

Private Sub cmdQuit_Click()
  Unload frmPrintQ
  End
End Sub

Private Sub Form_Load()
  Dim i As Integer
  ' Fill Print Queue combo
  '----------------------
  ccode% = NW_NWDrvrLoaded()
  If ccode% = 0 Then
    NetError.Show 1
    Exit Sub
  End If
  Dim NWBind As BinderyInfo
  cmbprintqueues.Clear
  ccode% = NW_ScanBindery(-1, &H3, NWBind)
  cmbprintqueues.AddItem NWBind.object_name
  
  Do
    ccode% = NW_ScanBindery(NWBind.object_ID, &H3, NWBind)
    If (ccode% = 0) Then
      cmbprintqueues.AddItem NWBind.object_name
    End If
  Loop Until ccode%
  
  ' Fill LPT combo
  '--------------------------
  cmbLPT.Clear
  For i = 1 To 9
    cmbLPT.AddItem i
  Next i
  
 
End Sub


