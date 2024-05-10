VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   4020
   ClientLeft      =   2220
   ClientTop       =   2055
   ClientWidth     =   8370
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
   Height          =   4395
   Left            =   2175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8370
   Top             =   1725
   Width           =   8460
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "List Objects With Equivalent Security"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1395
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Security Equivalent List:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "List The Security Equivalence of a Bindery Object 16-bit"
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
      Width           =   7695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   240
      X2              =   7920
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
    Dim cUserName As String * 48
    Dim cObjName As String * 48

    cUserName = LTrim$(RTrim$(Text1.TEXT)) + Chr$(0)
    
    List1.Clear

    For i = 0 To 31
        ccode% = NW_SecurityEquivalence(OT_USER, cUserName, i, cObjName)
        If ccode% = 0 Then
            List1.AddItem cObjName
        End If
    Next i
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

