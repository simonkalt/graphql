VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   3435
   ClientLeft      =   1275
   ClientTop       =   1950
   ClientWidth     =   8085
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
   Height          =   3810
   Left            =   1230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   8085
   Top             =   1620
   Width           =   8175
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Submit Job to Queue for Output"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
   End
   Begin VBX.NPrint NPrint1 
      AutoStart       =   0   'False
      CompletionCode  =   0
      Description     =   "NPRINT Submitted"
      EntryOpen       =   0   'False
      FIlename        =   "F:\PER\SRC\TEST.BIN"
      NumberofCopies  =   1
      ObjectType      =   3
      OperatorHold    =   -1  'True
      QueueName       =   "PRINTQ_10"
      ServiceRestart  =   0   'False
      UserHold        =   0   'False
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set a few properties and press the button..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"NPRINCTL.frx":0000
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
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   7800
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
    NPRINT1.Refresh
    Text1.TEXT = NPRINT1.CompletionCode
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

