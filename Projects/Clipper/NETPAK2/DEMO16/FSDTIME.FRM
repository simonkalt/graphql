VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   4410
   ClientLeft      =   1110
   ClientTop       =   1395
   ClientWidth     =   7470
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   4785
   Left            =   1065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7470
   Top             =   1065
   Width           =   7560
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   1095
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   " "
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      MaxLength       =   2
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set File Server Date/Time"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6600
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6600
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      MaxLength       =   2
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label12 
      Caption         =   "Day of Week:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   ":"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2040
      TabIndex        =   25
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   ":"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      TabIndex        =   24
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   ":"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6120
      TabIndex        =   20
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   ":"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4920
      TabIndex        =   19
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6120
      TabIndex        =   18
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4920
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "New File Server Date/Time:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current File Server Date:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2160
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   960
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get/Set File Server Date and Time - 16-bit"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   7095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   240
      X2              =   7320
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    Dim SetDT As SetFSDateTime
    Dim nYr As Integer
    Dim nMon As Integer
    Dim nDay As Integer
    Dim nHr As Integer
    Dim nMin As Integer
    Dim nSec As Integer


    nYr = Val(LTrim$(RTrim$(Text6.TEXT)))
    nMon = Val(LTrim$(RTrim$(Text4.TEXT)))
    nDay = Val(LTrim$(RTrim$(Text5.TEXT)))
    nHr = Val(LTrim$(RTrim$(Text7.TEXT)))
    nMin = Val(LTrim$(RTrim$(Text8.TEXT)))
    nSec = Val(LTrim$(RTrim$(Text9.TEXT)))

    SetDT.date_time(0) = nYr
    SetDT.date_time(1) = nMon
    SetDT.date_time(2) = nDay
    SetDT.date_time(3) = nHr
    SetDT.date_time(4) = nMin
    SetDT.date_time(5) = nSec
    
    ccode% = NW_SetFSDateTime(SetDT)
End Sub

Private Sub Command3_Click()
    
    Dim GFSDT As GetFSDateTime
    Dim DayOfWeek$(6)

    DayOfWeek$(0) = "Sunday"
    DayOfWeek$(1) = "Monday"
    DayOfWeek$(2) = "Tuesday"
    DayOfWeek$(3) = "Wednesday"
    DayOfWeek$(4) = "Thursday"
    DayOfWeek$(5) = "Friday"
    DayOfWeek$(6) = "Saturday"

    ccode% = NW_GetFSDateTime(GFSDT)

    Text1.TEXT = LTrim$(Str$(GFSDT.date_time(1)))
    Text2.TEXT = LTrim$(Str$(GFSDT.date_time(2)))
    Text3.TEXT = LTrim$(Str$(GFSDT.date_time(0)))

    Text10.TEXT = LTrim$(Str$(GFSDT.date_time(3)))
    Text11.TEXT = LTrim$(Str$(GFSDT.date_time(4)))
    Text12.TEXT = LTrim$(Str$(GFSDT.date_time(5)))

    Text13.TEXT = DayOfWeek$(GFSDT.date_time(6))


End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If
    
    Dim GFSDT As GetFSDateTime
    Dim DayOfWeek$(6)
    
    DayOfWeek$(0) = "Sunday"
    DayOfWeek$(1) = "Monday"
    DayOfWeek$(2) = "Tuesday"
    DayOfWeek$(3) = "Wednesday"
    DayOfWeek$(4) = "Thursday"
    DayOfWeek$(5) = "Friday"
    DayOfWeek$(6) = "Saturday"
    
    ccode% = NW_GetFSDateTime(GFSDT)

    Text1.TEXT = LTrim$(Str$(GFSDT.date_time(1)))
    Text2.TEXT = LTrim$(Str$(GFSDT.date_time(2)))
    Text3.TEXT = LTrim$(Str$(GFSDT.date_time(0)))

    Text10.TEXT = LTrim$(Str$(GFSDT.date_time(3)))
    Text11.TEXT = LTrim$(Str$(GFSDT.date_time(4)))
    Text12.TEXT = LTrim$(Str$(GFSDT.date_time(5)))
    
    Text13.TEXT = DayOfWeek$(GFSDT.date_time(6))

End Sub

