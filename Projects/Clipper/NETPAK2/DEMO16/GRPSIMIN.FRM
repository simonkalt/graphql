VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   4230
   ClientLeft      =   1380
   ClientTop       =   1755
   ClientWidth     =   8070
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
   Height          =   4605
   Left            =   1335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   8070
   Top             =   1425
   Width           =   8160
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1785
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Find List of Groups"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Group List:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "List The Groups A Bindery Object Is In 16-bit"
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7695
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
    Dim cUserName As String * 48
    Dim cObjName As String * 48

    cUserName = LTrim$(RTrim$(Text1.TEXT)) + Chr$(0)
    
    List1.Clear

    For i = 0 To 31
        ccode% = NW_GroupsImIn(OT_USER, cUserName, i, cObjName)
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

