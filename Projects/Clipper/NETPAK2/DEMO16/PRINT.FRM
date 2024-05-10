VERSION 4.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetPak Professional - Print Capture Flags"
   ClientHeight    =   6690
   ClientLeft      =   1590
   ClientTop       =   1590
   ClientWidth     =   7920
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Height          =   7065
   Left            =   1545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7920
   Top             =   1260
   Width           =   8010
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Set Flags"
      Height          =   375
      Left            =   540
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   540
      TabIndex        =   2
      Top             =   2010
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Get Flags"
      Height          =   375
      Left            =   540
      TabIndex        =   0
      Top             =   1110
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Netware Functions 16-bit"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   360
      TabIndex        =   51
      Top             =   180
      Width           =   5535
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2430
      TabIndex        =   50
      Top             =   2130
      Width           =   5175
   End
   Begin VB.Label Label49 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Banner User Name:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   540
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label48 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Job Number:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4380
      TabIndex        =   49
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label46 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Queue ID:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4740
      TabIndex        =   48
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Job Valid:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4680
      TabIndex        =   47
      Top             =   5880
      Width           =   1395
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Queue Flag:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4380
      TabIndex        =   46
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Capture In Progress:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4200
      TabIndex        =   45
      Top             =   3600
      Width           =   1905
   End
   Begin VB.Label Label42 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Connection ID Queue Print Job:"
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   4620
      TabIndex        =   44
      Top             =   4470
      Width           =   1455
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Timing Out Flag:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4560
      TabIndex        =   43
      Top             =   3960
      Width           =   1515
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "File Capture Flag:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4500
      TabIndex        =   42
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Currently Captured:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4380
      TabIndex        =   41
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Form Name:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4620
      TabIndex        =   40
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label37 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Maximum Characters:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   5880
      Width           =   1875
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Maximum Lines:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   840
      TabIndex        =   38
      Top             =   5520
      Width           =   1515
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Flush Capture On Device Close:"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   720
      TabIndex        =   37
      Top             =   5010
      Width           =   1635
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Timeout Count:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   960
      TabIndex        =   36
      Top             =   4680
      Width           =   1395
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Local LPT Device:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   660
      TabIndex        =   35
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Banner Text:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1200
      TabIndex        =   34
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Form Type:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1200
      TabIndex        =   33
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Copies:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   32
      Top             =   3225
      Width           =   1650
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Server Printer:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1020
      TabIndex        =   31
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tab Size:"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1140
      TabIndex        =   30
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   29
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   28
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   26
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   4650
      Width           =   495
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   23
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   21
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      X1              =   360
      X2              =   7560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Job Flags:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   360
      TabIndex        =   9
      Top             =   660
      Width           =   1935
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2430
      TabIndex        =   8
      Top             =   1770
      Width           =   5175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2430
      TabIndex        =   7
      Top             =   1410
      Width           =   5175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2430
      TabIndex        =   6
      Top             =   1050
      Width           =   5175
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2430
      TabIndex        =   5
      Top             =   690
      Width           =   5175
   End
End
Attribute VB_Name = "Main"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Const RELEASE_JOB = 4
    Const SUPPRESS_FORM_FEED = 8
    Const NOTIFY = 16
    Const INTERPRET_CODES = 64
    Const PRINT_BANNER = 128

    Dim Getflags As GetCaptureFlags

    '--- function call
    ccode% = NW_GetDefCapFlags(Getflags)

    If (ccode% = 0) Then
        '--- evaluate flag field bit definitions
        If (Asc(Getflags.flags) And RELEASE_JOB_NOT_CONN) Then
            Label3.Caption = "Release print job if connection is lost."
        Else
            Label3.Caption = "Do not release print job if connection is lost."
        End If
            
        If (Asc(Getflags.flags) And SUPPRESS_FORM_FEED) Then
            Label4.Caption = "Suppress form feed at the end of the job."
        Else
            Label4.Caption = "Do not suppress form feed at the end of the job."
        End If

        If (Asc(Getflags.flags) And NOTIFY) Then
            Label1.Caption = "Notify user when job is printed."
        Else
            Label1.Caption = "Do not notify user when job is printed."
        End If
        
        If (Asc(Getflags.flags) And INTERPRET_CODES) Then
            Label5.Caption = "Print service will interpret printer control sequences."
        Else
            Label5.Caption = "Print service will not interpret printer control sequences."
        End If

        If (Asc(Getflags.flags) And PRINT_BANNER) Then
            Label6.Caption = "Print banner page."
        Else
            Label6.Caption = "Do not print banner page."
        End If
        
        '--- tab size
        Label8.Caption = Str$(Asc(Getflags.tab_size))

        '--- server printer
        Label9.Caption = Str$(Asc(Getflags.server_printer))

        '--- number of copies
        Label10.Caption = Str$(Asc(Getflags.number_of_copies))

        '--- form type
        Label11.Caption = Str$(Asc(Getflags.form_type))
    
        '--- banner text
        Label12.Caption = Getflags.banner_text
    
        '--- local LPT device
        Label13.Caption = Str$(Asc(Getflags.local_lpt_device))

        '--- flush capture timeout
        Label14.Caption = Str$(Getflags.flush_capture_timeout_count)
    
        '--- flush capture on device close
    
        If (Asc(Getflags.flush_capture_on_device_close) = 0) Then
            Label15.Caption = "Enabled"
        Else
            Label15.Caption = "Diabled"
        End If
    
        '--- maximum lines
        Label16.Caption = Getflags.max_lines

        '--- maximum characters
        Label17.Caption = Getflags.max_chars

        '--- form name
        Label18.Caption = Getflags.form_name

        '--- LPT capture flag
        If (Asc(Getflags.lpt_capture_flag) = 0) Then
            Label19.Caption = "No"
        Else
            Label19.Caption = "Yes"
        End If

        '--- File capture flag
        If (Asc(Getflags.file_capture_flag) = 0) Then
            Label20.Caption = "No"
        Else
            Label20.Caption = "Yes"
        End If

        '--- Timing out flag
        If (Asc(Getflags.timing_out_flag) = 0) Then
            Label21.Caption = "No"
        Else
            Label21.Caption = "Yes"
        End If

        '--- connection ID queue print job
        Label22.Caption = Str$(Asc(Getflags.connection_id_queue_print_job))
    
        '--- capture in progress
        If (Asc(Getflags.capture_in_progress) = 0) Then
            Label23.Caption = "No"
        Else
            Label23.Caption = "Yes"
        End If

        If (Asc(Getflags.print_queue_flag) = 255) Then
            Label24.Caption = "Yes"
        Else
            Label24.Caption = "No"
        End If

        If (Asc(Getflags.print_job_valid) = 0) Then
            Label25.Caption = "No"
        Else
            Label25.Caption = "Yes"
        End If
    
        If (Asc(Getflags.print_queue_flag) = 255) Then
            Label26.Caption = Hex$(Getflags.print_queue_id)
        Else
            Label26.Caption = ""
        End If

        Label27.Caption = Str$(Getflags.print_job_number)

        User$ = Space$(13)
        
        ccode% = NW_GetBannerUName(User$)

        Label48.Caption = User$
    Else
        '--- display completion code message
        Response% = MsgBox("Error: " + Hex$(ccode%), 64, "NW_GetDefCapFlags Function Call")
    End If
End Sub

Private Sub Command2_Click()
    
    Const RELEASE_JOB = 4
    Const SUPPRESS_FORM_FEED = 8
    Const NOTIFY = 16
    Const INTERPRET_CODES = 64
    Const PRINT_BANNER = 128

    Dim Setflags As SetCaptureFlags
    
    '---- End any existing capture
    ccode% = NW_EndLPTCap%()

    '---- set the capture print queue
    ccode% = NW_SetCapturePrintQ%(0, "PQ_01")

    '---- set default capture flags
    Setflags.status = Chr$(0)
    
    flags = flags + RELEASE_JOB
    flags = flags + SUPPRESS_FORM_FEED
    flags = flags + NOTIFY
    flags = flags + INTERPRET_CODES
    flags = flags + PRINT_BANNER
    Setflags.flags = Chr$(flags)
    
    Setflags.tab_size = Chr$(8)
    
    Setflags.number_of_copies = Chr$(1)

    Setflags.banner_text = "CS Banner"
    
    Setflags.max_lines = NW_SwapInteger(66)
    Setflags.max_chars = NW_SwapInteger(132)

    ccode% = NW_SetDefCapFlags%(Setflags)
    
    '---- start LPT capture
    ccode% = NW_StartLPTCap%()
    
End Sub

Private Sub Command3_Click()
   
   Form_Unload False

End Sub

Private Sub Form_Load()
    ccode% = NW_NWDrvrLoaded()
    If ccode% = 0 Then
        NetError.Show 1
    End If
    
    Main.Left = Screen.Width \ 2 - Main.Width \ 2
    Main.TOP = Screen.Height \ 2 - Main.Height \ 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

