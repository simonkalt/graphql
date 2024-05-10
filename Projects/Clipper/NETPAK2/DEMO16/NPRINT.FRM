VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   4845
   ClientLeft      =   1695
   ClientTop       =   2505
   ClientWidth     =   7560
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
   Height          =   5220
   Left            =   1650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7560
   Top             =   2175
   Width           =   7650
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   " "
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Submit File To Queue For Ouput"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Enter Form Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Enter File Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter Queue Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"NPRINT.frx":0000
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
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   7335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   7440
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
    ccode% = NW_NPRINT(OT_PRINT_QUEUE, LTrim$(RTrim$(Text1.TEXT)), LTrim$(RTrim$(Text3.TEXT)), 0, 0, "NPRINT Submitted", LTrim$(RTrim$(Text4.TEXT)))
    Text2.TEXT = Hex$(ccode%)
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

