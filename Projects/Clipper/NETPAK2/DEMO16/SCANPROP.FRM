VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetPak Professional Demonstration"
   ClientHeight    =   4155
   ClientLeft      =   1290
   ClientTop       =   1935
   ClientWidth     =   7425
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
   Height          =   4530
   Left            =   1245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7425
   Top             =   1605
   Width           =   7515
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Scan For Properties"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1785
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Property List:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bindery Object Name:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scan For Bindery Object Properties 16-bit"
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
      TabIndex        =   3
      Top             =   360
      Width           =   7215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   7200
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
    Dim NWPropInfo As PropertyInfo
    Dim cUserName As String * 48

    cUserName = LTrim$(RTrim$(Text1.TEXT)) + Chr$(0)

    List1.Clear
    ccode% = NW_ScanProperty(OT_USER, cUserName, -1, NWPropInfo)
    List1.AddItem NWPropInfo.property_name
    
    Do
        ccode% = NW_ScanProperty(OT_USER, cUserName, NWPropInfo.sequence_number, NWPropInfo)
        If (ccode% = 0) Then
            List1.AddItem NWPropInfo.property_name
        End If
    Loop Until ccode%
End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If

End Sub

