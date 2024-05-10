VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   4140
   ClientLeft      =   2070
   ClientTop       =   2190
   ClientWidth     =   6450
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
   Height          =   4515
   Left            =   2025
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6450
   Top             =   1860
   Width           =   6540
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Obtain Current Job List"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   1590
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Jobs Currently In Queue:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Queue Job List:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter Queue Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "List The Jobs Currently In A Queue 16-bit"
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
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   240
      X2              =   5040
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
    Dim Jobs As QueueJobs
    ccode% = NW_GetQueueJobList(OT_PRINT_QUEUE, LTrim$(RTrim$(Text1.TEXT)), Jobs)

    List1.Clear
    For i = 0 To Jobs.job_count - 1
        List1.AddItem LTrim$(RTrim$(Str$(Jobs.job_list(i))))
    Next
    
    Text2.TEXT = Hex$(ccode%)
    Text3.TEXT = Jobs.job_count
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

