VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   3360
   ClientLeft      =   825
   ClientTop       =   2175
   ClientWidth     =   5985
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
   Height          =   3735
   Left            =   780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5985
   Top             =   1845
   Width           =   6075
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Destroy Bindery Object (use with caution!)"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Bindery Object Type:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Completion Code:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Destroy A Bindery Object 16-bit"
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
      TabIndex        =   4
      Top             =   360
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   240
      X2              =   5640
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

    nReturn% = MsgBox("Are you sure that you want to delete the object " + LTrim$(RTrim$(Text1.TEXT)), vbYesNo, "DELETE BINDERY OBJECT !!!")
    
    If nReturn% = 6 Then
        ccode% = NW_DeleteObject(OT_USER, LTrim$(RTrim$(Text1.TEXT)))
    Else
        ccode% = &HFF
    End If
    
    Text2.TEXT = Hex$(ccode%)
    
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

