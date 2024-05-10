Attribute VB_Name = "Module7"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' FILENAME:     nwdecl16.bas
'
' DATE:         08/02/95
'
' VERSION:      2.00.000
'
' DESCRIPTION:  This module contains all the NetPak helper functions,
'                   structure declarations, and constants needed for use
'                   with the 16-bit npnw100 and npprowfw DLL's.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Bindery security definitions
Global Const BS_ANY_READ = &H0                     'Readable by anyone
Global Const BS_LOGGED_READ = &H1                  'Must be logged in to read
Global Const BS_OBJECT_READ = &H2                  'Readable by same object or supervisor
Global Const BS_SUPER_READ = &H3                   'Readable by supervisor only
Global Const BS_BINDERY_READ = &H4                 'Readable only by the bindery
Global Const BS_ANY_WRITE = &H0                    'Writeable by anyone
Global Const BS_LOGGED_WRITE = &H10                'Must be logged in to write
Global Const BS_OBJECT_WRITE = &H20                'Writeable by same object or supervisor
Global Const BS_SUPER_WRITE = &H30                 'Writeable only by the supervisor
Global Const BS_BINDERY_WRITE = &H40               'Writeable by the bindery only

'Bindery object type definitions
Global Const OT_WILD = (-1)                        'Matches any type
Global Const OT_UNKNOWN = &H0                      'Unknown object type
Global Const OT_USER = &H1                         'User
Global Const OT_USER_GROUP = &H2                   'Group
Global Const OT_PRINT_QUEUE = &H3                  'Print queue
Global Const OT_FILE_SERVER = &H4                  'File servers
Global Const OT_JOB_SERVER = &H5                   'Job server
Global Const OT_GATEWAY = &H6                      'Gateway
Global Const OT_PRINT_SERVER = &H7                 'Print server
Global Const OT_ARCHIVE_QUEUE = &H8                'Archive queue
Global Const OT_ARCHIVE_SERVER = &H9               'Archive server
Global Const OT_JOB_QUEUE = &HA                    'Job queue
Global Const OT_ADMINISTRATION = &HB               'Administration
Global Const OT_NAS_SNA_GATEWAY = &H21             'NAS SNA Gateway
Global Const OT_REMOTE_BRIDGE_SERVER = &H24        'Remote bridge server
Global Const OT_TIME_SYNCHRONIZATION_SERVER = &H2D 'Time Sync Server
Global Const OT_ARCHIVE_SERVER_DYNAMIC_SAP = &H2E  'Archive server dynamic SAP
Global Const OT_ADVERTISING_PRINT_SERVER = &H47    'Advertising print server
Global Const OT_BTRIEVE_VAP = &H4B                 'Btrieve VAP
Global Const OT_PRINT_QUEUE_USER = &H53            'Print queue user

'Trustee Access Rights in a network directory
Global Const TA_NONE = &H0
Global Const TA_READ = &H1
Global Const TA_WRITE = &H2
Global Const TA_OPEN = &H4
Global Const TA_CREATE = &H8
Global Const TA_DELETE = &H10
Global Const TA_OWNERSHIP = &H20
Global Const TA_SEARCH = &H40
Global Const TA_MODIFY = &H80
Global Const TA_ALL = &HFF

'Attributes of objects and properties in the bindery
Global Const BF_STATIC = &H0
Global Const BF_DYNAMIC = &H1
Global Const BF_ITEM = &H0
Global Const BF_SET = &H2

'Queue Job Control Flags
Global Const QF_OPERATOR_HOLD = &H80
Global Const QF_USER_HOLD = &H40
Global Const QF_ENTRY_OPEN = &H20
Global Const QF_SERVICE_RESTART = &H10
Global Const QF_AUTO_START = &H8

Public Declare Function NW_NWDrvrLoaded% Lib "NPNW16.DLL" ()
Public Declare Function NW_SwapInteger% Lib "NPNW16.DLL" (ByVal nNbr As Integer)
Public Declare Function NW_SwapLong& Lib "NPNW16.DLL" (ByVal nNbr As Long)
Public Declare Function NW_WWGDrvrLoaded% Lib "NPNW16.DLL" ()

'----------------------------------------
' Netware Bindery Services
'----------------------------------------
Type BinderyInfo
    object_ID As Long
    object_type As Integer
    object_name As String * 48
    object_flag As String * 1
    object_security As String * 1
    object_has_properties As String * 1
End Type

Type PropertyInfo
    property_name As String * 16
    property_flags As String * 1
    property_security As String * 1
    sequence_number As Long
    property_has_value As String * 1
    more_properties As String * 1
End Type

Type ReadLoginControl
    aexpir_date_year As String * 1
    aexpir_date_month As String * 1
    aexpir_date_day As String * 1
    dis_flag As String * 1
    pexpir_date_year As String * 1
    pexpir_date_month As String * 1
    pexpir_date_day As String * 1
    grace_logins As String * 1
    pexpir_int As Integer
    grace_reset_value As String * 1
    min_pwd_length As String * 1
    max_concurrent_conn As Integer
    last_login_date_year As String * 1
    last_login_date_month As String * 1
    last_login_date_day As String * 1
    last_login_date_hour As String * 1
    last_login_date_minute As String * 1
    last_login_date_second As String * 1
    restrictions_flag As String * 1
    unused As String * 1
    max_disk_usage As Long
    bad_login_count As Integer
    next_reset_time As Long
End Type

Type WriteLoginControl
    aexpir_date_year As String * 1
    aexpir_date_month As String * 1
    aexpir_date_day As String * 1
    dis_flag As String * 1
    pexpir_date_year As String * 1
    pexpir_date_month As String * 1
    pexpir_date_day As String * 1
    grace_logins As String * 1
    pexpir_int As Integer
    grace_reset_value As String * 1
    min_pwd_length As String * 1
    max_concurrent_conn As Integer
End Type

Type AllowedTimes
    allowed_login_time(42) As String * 1
End Type

Public Declare Function NW_AddBObjToSet% Lib "NPNW16.DLL" (ByVal nOType As Integer, ByVal cObjName$, ByVal cPropName$, ByVal nMType As Integer, ByVal cMemName$)
Public Declare Function NW_ChangeBObjPwd% Lib "NPNW16.DLL" (ByVal cObjName$, ByVal cOldPwd$, ByVal cNewPwd$)
Public Declare Function NW_ChangeBObjSecurity% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, ByVal nNewSecurity As Integer)
Public Declare Function NW_ChangePropSecurity% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, ByVal cPropName$, ByVal nNewSecurity As Integer)
Public Declare Function NW_CloseBindery% Lib "NPNW16.DLL" ()
Public Declare Function NW_CreateObject% Lib "NPNW16.DLL" (ByVal nFlag As Integer, ByVal nSecurity As Integer, ByVal nType As Integer, ByVal cOName$)
Public Declare Function NW_CreateProperty% Lib "NPNW16.DLL" (ByVal nType As Integer, ByVal cOName$, ByVal nFlags As Integer, ByVal nSecurity As Integer, ByVal cPName$)
Public Declare Function NW_DelBObjFromSet% Lib "NPNW16.DLL" (ByVal nOType As Integer, ByVal cObjName$, ByVal cPropName$, ByVal nMType As Integer, ByVal cMemName$)
Public Declare Function NW_DeleteObject% Lib "NPNW16.DLL" (ByVal nType%, ByVal cOName$)
Public Declare Function NW_DeleteProperty% Lib "NPNW16.DLL" (ByVal nType%, ByVal cOName$, ByVal cPName$)
Public Declare Function NW_GetBAccessLvl% Lib "NPNW16.DLL" (nSecLevel%)
Public Declare Function NW_GetBObjectID% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, nObjID&)
Public Declare Function NW_GetBObjName% Lib "NPNW16.DLL" (ByVal nObjID As Long, ByVal nObjName$)
Public Declare Function NW_GroupMembers% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjNameIn$, ByVal nSeqNbr As Integer, ByVal cObjNameOut$)
Public Declare Function NW_GroupsImIn% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjNameIn$, ByVal nSeqNbr As Integer, ByVal cObjNameOut$)
Public Declare Function NW_IsObjInSet% Lib "NPNW16.DLL" (ByVal nOType As Integer, ByVal cObjName$, ByVal cPropName$, ByVal nMType As Integer, ByVal cMemName$)
Public Declare Function NW_OpenBindery% Lib "NPNW16.DLL" ()
Public Declare Function NW_ReadLoginControl% Lib "NPNW16.DLL" (ByVal cUName$, LC As ReadLoginControl, AT As AllowedTimes)
Public Declare Function NW_ReadPropertyValue% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cOName$, ByVal cPropName$, ByVal cPropValue$)
Public Declare Function NW_ScanBindery% Lib "NPNW16.DLL" (ByVal nLast As Long, ByVal nObjType As Integer, NWBindInfo As BinderyInfo)
Public Declare Function NW_ScanProperty% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, ByVal nSeqNbr As Long, NWPropInfo As PropertyInfo)
Public Declare Function NW_SecurityEquivalence% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjNameIn$, ByVal nSeqNbr As Integer, ByVal cObjNameOut$)
Public Declare Function NW_UserFullName% Lib "NPNW16.DLL" (ByVal cUName$, ByVal cPropVal$)
Public Declare Function NW_VerifyBObjPwd% Lib "NPNW16.DLL" (ByVal cObjName$, ByVal cPwd$)
Public Declare Function NW_WriteLoginControl% Lib "NPNW16.DLL" (ByVal cUName$, LC As WriteLoginControl, AT As AllowedTimes)
Public Declare Function NW_WritePropertyValue% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cOName$, ByVal cPropName$, ByVal cPropValue$)
Public Declare Function NW_RenameBObj% Lib "NPNW16.DLL" (ByVal nObjType%, ByVal cNewObjName$, ByVal cOldObjName)
Public Declare Function NW_Surname% Lib "NPNW16.DLL" (ByVal cUserName$, ByVal cPropertyValue$)

'----------------------------------------
' Netware Connection Services
'----------------------------------------
Type List
    connection_list(101) As Integer
End Type

Public Declare Function NW_AttachToFS% Lib "NPNW16.DLL" (ByVal cFSName$, nConnID%)
Public Declare Function NW_DetachFromFS% Lib "NPNW16.DLL" (ByVal nConnID%)
Public Declare Function NW_EnterLoginArea% Lib "NPNW16.DLL" (ByVal cPath$, ByVal nNbrDrives As Integer)
Public Declare Function NW_GetConnInfo% Lib "NPNW16.DLL" (ByVal nConnNbr%, ByVal cUserName$)
Public Declare Function NW_GetConnInfoF2% Lib "NPNW16.DLL" (ByVal nConnNbr%, ByVal cUserName$)
Public Declare Function NW_GetConnNbr% Lib "NPNW16.DLL" ()
Public Declare Function NW_GetObjConnNbrs% Lib "NPNW16.DLL" (ByVal cUName$, nNbrConns%, ConnList As List)
Public Declare Function NW_LoginToFS% Lib "NPNW16.DLL" (ByVal cUserName$, ByVal cUserPwd$)
Public Declare Function NW_Logout% Lib "NPNW16.DLL" ()
Public Declare Function NW_LogoutFromFS% Lib "NPNW16.DLL" (ByVal nConnID As Integer)
Public Declare Function NW_GetInternetAddress% Lib "NPNW16.DLL" (ByVal nConnNbr%, ByVal cAddress$)
Public Declare Function NW_GetStationAddress% Lib "NPNW16.DLL" (ByVal cNodeAddress$)

'----------------------------------------
' Netware Directory Services
'----------------------------------------
Type VolumeInfo
    total_blocks As Long
    free_blocks As Long
    purgeable_blocks As Long
    not_yet_purgeable_blocks As Long
    total_directory_entries As Long
    available_directory_entries As Long
    flags As Long
    sectors_per_block As String * 1
    volume_name_length As String * 1
    volume_name As String * 48
End Type

Type Trustees
    number_entries As String * 1
    object_ID(20) As Long
    rights(20) As Integer
End Type

Public Declare Function NW_AddTrusteeToDirF2% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, ByVal nDskNbr As Integer, ByVal cPath$)
Public Declare Function NW_CreateDir% Lib "NPNW16.DLL" (ByVal nDskNbr As Integer, ByVal nTRights As Integer, ByVal cPath$)
Public Declare Function NW_DeleteDir% Lib "NPNW16.DLL" (ByVal nDskNbr As Integer, ByVal cPath$)
Public Declare Function NW_DelTrusteeFromDir% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, ByVal nDskNbr As Integer, ByVal cPath$)
Public Declare Function NW_GetDirPath% Lib "NPNW16.DLL" (ByVal nDskNbr As Integer, ByVal cPath$)
Public Declare Function NW_GetVolInfoF2% Lib "NPNW16.DLL" (ByVal nVolNbr As Integer, VolInfo As VolumeInfo)
Public Declare Function NW_GetVolumeName% Lib "NPNW16.DLL" (ByVal nVolNbr As Integer, ByVal cVolName$)
Public Declare Function NW_GetVolumeNbr% Lib "NPNW16.DLL" (ByVal cVolName As String, nVolNbr%)
Public Declare Function NW_MapFakeRoot% Lib "NPNW16.DLL" (ByVal cDrvrLtr$, ByVal cPathName$)
Public Declare Function NW_RenameDir% Lib "NPNW16.DLL" (ByVal cOldDir$, ByVal cNewDir$)
Public Declare Function NW_ScanEntryforTteesF2% Lib "NPNW16.DLL" (ByVal nDskNbr As Integer, ByVal cPath$, ScanTtees As Trustees)
Public Declare Function NW_ScanTrusteePathF2% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal nVolNbr As Integer, ByVal cObjName$, ByVal nSeqNbr As Integer, ByVal cPath$, nNext%)
Public Declare Function NW_MapDirNumberToPath% Lib "NPNW16.DLL" (ByVal nDirHandle&, ByVal cPath$)
Public Declare Function NW_ScanFileForTteeF2% Lib "NPNW16.DLL" (ByVal nDskNbr%, ByVal cPath$, ST As Trustees)

'-----------------------------------------
' Netware File Server Environment Services
'-----------------------------------------
Type FileServerInfo
    server_name As String * 48
    netware_version As String * 1
    netware_subversion As String * 1
    connections_supported As Integer
    connections_in_use As Integer
    max_connected_volumes As Integer
    os_revision_number As String * 1
    sft_level As String * 1
    tts_level As String * 1
    peak_connections_used As Integer
    accounting_version As String * 1
    vap_version As String * 1
    queuing_version As String * 1
    print_server_version As String * 1
    virtual_console_version As String * 1
    security_restrictions_level As String * 1
    internet_bridge_version As String * 1
End Type

Type FServerInfo
    server_name As String * 48
    netware_version As String * 1
    netware_subversion As String * 1
    connections_supported As Integer
    connections_in_use As Integer
    max_connected_volumes As Integer
    os_revision_number As String * 1
    sft_level As String * 1
    tts_level As String * 1
    peak_connections_used As Integer
    accounting_version As String * 1
    vap_version As String * 1
    queuing_version As String * 1
    print_server_version As String * 1
    virtual_console_version As String * 1
    security_restrictions_level As String * 1
    internet_bridge_version As String * 1
    reserved As String * 60
End Type

Type ScanCnUsingFile
    connection_number(70) As Integer
    task_number(70) As Integer
End Type

Type GetFSDateTime
    date_time(7) As Integer
End Type

Type SetFSDateTime
    date_time(6) As Integer
End Type

Public Declare Function NW_BcastConsole% Lib "NPNW16.DLL" (ByVal cMsg$)
Public Declare Function NW_ChkConsolePriv% Lib "NPNW16.DLL" ()
Public Declare Function NW_ClearConnNbr% Lib "NPNW16.DLL" (ByVal nConnNbr As Integer)
Public Declare Function NW_ClearConnNbrF2% Lib "NPNW16.DLL" (ByVal nConnNbr As Integer)
Public Declare Function NW_DisableFSLogin% Lib "NPNW16.DLL" ()
Public Declare Function NW_DisableTTS% Lib "NPNW16.DLL" ()
Public Declare Function NW_DownFS% Lib "NPNW16.DLL" (ByVal nForceFlag As Integer)
Public Declare Function NW_EnableFSLogin% Lib "NPNW16.DLL" ()
Public Declare Function NW_EnableTTS% Lib "NPNW16.DLL" ()
Public Declare Function NW_FSFileCopy% Lib "NPNW16.DLL" (ByVal cSource$, ByVal cDest$)
Public Declare Function NW_FSSerialNbr% Lib "NPNW16.DLL" (ByVal cSerialNumber$, nAppNbr As Integer)
Public Declare Function NW_GetFileServerInfo% Lib "NPNW16.DLL" (FSInfo As FileServerInfo)
Public Declare Function NW_GetFileServerInfoF2% Lib "NPNW16.DLL" (FSvrInfo As FServerInfo)
Public Declare Function NW_GetFSDateTime% Lib "NPNW16.DLL" (GFSDT As GetFSDateTime)
Public Declare Function NW_GetFSDescStrs% Lib "NPNW16.DLL" (ByVal cDescStrings$)
Public Declare Function NW_GetSpaceUsedF2% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, ByVal nVolNbr As Integer, nInUse&, nRestrict&)
Public Declare Function NW_ScanConnUsingAFileF2% Lib "NPNW16.DLL" (ByVal cPath$, nCount%, ScanUsing As ScanCnUsingFile)
Public Declare Function NW_SetFSDateTime% Lib "NPNW16.DLL" (SFSDT As SetFSDateTime)
Public Declare Function NW_AddUsrSpaceRestF2% Lib "NPNW16.DLL" (ByVal cObjName$, ByVal nVolNbr As Integer, ByVal nRestrict As Integer)
Public Declare Function NW_GetConnOpFilesF2% Lib "NPNW16.DLL" (ByVal nConnNbr As Integer, ByVal nRecord As Integer, ByRef nRecords As Integer, ByRef nVolNbr As Integer, ByRef nParentDir As Long, ByRef cFile$)

'----------------------------------------
' Netware Message Services
'----------------------------------------
Public Declare Function NW_GetBroadcastMessage% Lib "NPNW16.DLL" (ByVal cMessage$)
Public Declare Function NW_GetBroadcastMode% Lib "NPNW16.DLL" ()
Public Declare Function NW_LogNetMsg% Lib "NPNW16.DLL" (ByVal cMsg$)
Public Declare Function NW_SendBroadcastMessage% Lib "NPNW16.DLL" (ByVal cMessage As String, ByVal nConnNbr As Integer)
Public Declare Function NW_SetBroadcastMode% Lib "NPNW16.DLL" (ByVal nMode As Integer)

'----------------------------------------
' Netware Print Services
'----------------------------------------
' you can read 63 bytes
Type GetCaptureFlags
    status As String * 1
    flags As String * 1
    tab_size As String * 1
    server_printer As String * 1
    number_of_copies As String * 1
    form_type As String * 1
    reserved_1 As String * 1
    banner_text As String * 13
    reserved_2 As String * 1
    local_lpt_device As String * 1
    flush_capture_timeout_count As Integer
    flush_capture_on_device_close As String * 1
    max_lines As Integer
    max_chars As Integer
    form_name As String * 13
    lpt_capture_flag As String * 1
    file_capture_flag As String * 1
    timing_out_flag As String * 1
    printer_setup_buffer_address As Long
    printer_reset_buffer_address As Long
    connection_id_queue_print_job As String * 1
    capture_in_progress As String * 1
    print_queue_flag As String * 1
    print_job_valid As String * 1
    print_queue_id As Long
    print_job_number As Integer
End Type

' you can change 42 bytes
Type SetCaptureFlags
    status As String * 1
    flags As String * 1
    tab_size As String * 1
    server_printer As String * 1
    number_of_copies As String * 1
    form_type As String * 1
    reserved_1 As String * 1
    banner_text As String * 13
    reserved_2 As String * 1
    local_lpt_device As String * 1
    flush_capture_timeout_count As Integer
    flush_capture_on_device_close As String * 1
    max_lines As Integer
    max_chars As Integer
    form_name As String * 13
End Type

Public Declare Function NW_CanLPTCap% Lib "NPNW16.DLL" ()
Public Declare Function NW_CanSPLPTCap% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer)
Public Declare Function NW_EndLPTCap% Lib "NPNW16.DLL" ()
Public Declare Function NW_EndSPLPTCap% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer)
Public Declare Function NW_FlushLPTCap% Lib "NPNW16.DLL" ()
Public Declare Function NW_FlushSPLPTCap% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer)
Public Declare Function NW_GetBannerUName% Lib "NPNW16.DLL" (ByVal cUName$)
Public Declare Function NW_GetDefCapFlags% Lib "NPNW16.DLL" (GetCFlags As GetCaptureFlags)
Public Declare Function NW_GetDefLclPtr% Lib "NPNW16.DLL" ()
Public Declare Function NW_GetLPTCapStat% Lib "NPNW16.DLL" (nConnID%)
Public Declare Function NW_GetSPCapFlags% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer, GetCFlags As GetCaptureFlags)
Public Declare Function NW_NPRINT% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal cFName$, ByVal nNbrCopies As Integer, ByVal nFlags As Integer, ByVal cDesc$, ByVal cFormName$)
Public Declare Function NW_SetBannerUName% Lib "NPNW16.DLL" (ByVal UName$)
Public Declare Function NW_SetCapturePrintQ% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer, ByVal cQName$)
Public Declare Function NW_SetDefLclPtr% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer)
Public Declare Function NW_SetDefCapFlags% Lib "NPNW16.DLL" (SetCFlags As SetCaptureFlags)
Public Declare Function NW_SetSPCapFlags% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer, SetCFlags As SetCaptureFlags)
Public Declare Function NW_SpecifyCapFile% Lib "NPNW16.DLL" (ByVal nDirHandle As Integer, ByVal cPath$)
Public Declare Function NW_StartLPTCap% Lib "NPNW16.DLL" ()
Public Declare Function NW_StartSPLPTCap% Lib "NPNW16.DLL" (ByVal nLPTDev As Integer)

'----------------------------------------
' Netware Queue Services
'----------------------------------------
Type QueueJobs
    job_count As Integer
    job_list(250) As Integer
    max_job_numbers As Integer
End Type

Type ReadQueueJobEntry
    client_station As String * 1
    client_task_num As String * 1
    client_id_num As Long
    target_server_id_num As Long
    target_execution_time As String * 6
    job_entry_time As String * 6
    job_number As Integer
    job_type As Integer
    job_position As String * 1
    job_control_flags As String * 1
    job_file_name As String * 14
    job_file_handle As String * 6
    server_station As String * 1
    server_task_number As String * 1
    server_id_number As Long
    text_job_description As String * 50
    client_record_area As String * 152
End Type

Type ChgQueueJobEntry
    client_station As String * 1
    client_task_num As String * 1
    client_id_num As Long
    target_execution_time As String * 6
    job_entry_time As String * 6
    job_type As Integer
    job_position As String * 1
    job_control_flags As String * 1
    job_file_name As String * 14
    job_file_handle As String * 6
    server_station As String * 1
    server_task_number As String * 1
    server_id_number As Long
    text_job_description As String * 50
    client_record_area As String * 152
End Type

Type QueueStatus
    queue_id As Long
    queue_status As String * 1
    number_of_jobs As String * 1
    number_of_servers As String * 1
    server_id_list(24) As Long
    server_stations_list As String * 25
    max_number_of_servers As String * 1
End Type

Public Declare Function NW_AbortSvcQJob% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal nJobNbr As Integer)
Public Declare Function NW_AttachQServerToQ% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$)
Public Declare Function NW_ChangeQJobEntry% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cObjName$, ByVal nJobNbr As Integer, CQJob As ChgQueueJobEntry)
Public Declare Function NW_ChangeQPos% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal nJobNbr As Integer, ByVal nPosition As Integer)
Public Declare Function NW_ChangetoCltRights% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal nJobNbr As Integer)
Public Declare Function NW_CloseFileStartQJob% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal nJobNbr As Integer)
Public Declare Function NW_CreateQueue% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$)
Public Declare Function NW_DestroyQueue% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$)
Public Declare Function NW_DetachQServerFrQ% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$)
Public Declare Function NW_FinishSvcQJobFile% Lib "NPNW16.DLL" (ByVal nType As Integer, ByVal cQName$, ByVal nJobNbr As Integer, ByVal nCharge As Long)
Public Declare Function NW_GetQJobFileSize% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal nJobNbr As Integer, nSize&)
Public Declare Function NW_GetQueueJobList% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, QJobs As QueueJobs)
Public Declare Function NW_ReadQJobEntry% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal nJobNbr As Integer, ReadQJob As ReadQueueJobEntry)
Public Declare Function NW_ReadQCurrentStatus% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ReadQStat As QueueStatus)
Public Declare Function NW_RemoveJob% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByVal nJobNbr As Integer)
Public Declare Function NW_SetQueueCurrentStatus% Lib "NPNW16.DLL" (ByVal nObjType As Integer, ByVal cQName$, ByRef nQStatus As Long)

'-------------------------------
'NetWare TTS Functions Functions
'-------------------------------
Public Declare Function NW_TTSAbort% Lib "NPNW16.DLL" ()
Public Declare Function NW_TTSAvailable% Lib "NPNW16.DLL" ()
Public Declare Function NW_TTSBegin% Lib "NPNW16.DLL" ()
Public Declare Function NW_TTSEnd% Lib "NPNW16.DLL" (nTransNbr%)
Public Declare Function NW_TTSStatus% Lib "NPNW16.DLL" (ByVal nTransNbr As Integer)

'----------------------------------------
' Netware Workstation Services
'----------------------------------------
Type ConnectionIDTable
  inUse As String * 1
  order As String * 1
  networkNumber As String * 4
  networkNode As String * 6
  socket As Integer
  receiveTimeOut As Integer
  routerNode As String * 6
  packetSequence As String * 1
  connectionNumber As String * 1
  connectionStatus As String * 1
  maximumTimeOut As Integer
  connectionWord As Integer
  majorServerVersion As String * 1
  serverFlags As String * 1
  minorServerVersion As String * 1
End Type

Type FileServerNameTable
  file_server As String * 48
End Type

Public Declare Function NW_GetConnID% Lib "NPNW16.DLL" (ByVal nIndex%, ConnTable As ConnectionIDTable)
Public Declare Function NW_GetDefConnID% Lib "NPNW16.DLL" ()
Public Declare Function NW_GetFileServerName% Lib "NPNW16.DLL" (ByVal nIndex%, FSNTable As FileServerNameTable)
Public Declare Function NW_GetPrefConnID% Lib "NPNW16.DLL" ()
Public Declare Function NW_GetPrmConnID% Lib "NPNW16.DLL" ()
Public Declare Function NW_GetNWShellVer% Lib "NPNW16.DLL" (ByVal cVersion$, nType%)
Public Declare Function NW_SetPrefConnID% Lib "NPNW16.DLL" (ByVal nConnID As Integer)
Public Declare Function NW_SetPrmConnID% Lib "NPNW16.DLL" (ByVal nConnID As Integer)

'-------------------------------
'Windows Functions
'-------------------------------

Public Declare Function WFW_NetConnectDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal nType As Integer)
Public Declare Function WFW_NetConnectionDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal nType As Integer)
Public Declare Function WFW_NetDeviceMode% Lib "NPPROW16.DLL" (ByVal nhwnd%)
Public Declare Function WFW_NetDisconnectDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal nType As Integer)
Public Declare Function WFW_NetBrowseDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal nType As Integer, ByVal cPath$)
Public Declare Function WFW_NetCancelConnection% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal cDrive$)
Public Declare Function WFW_NetAddConnection% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal cPath$, ByVal cPwd$, ByVal cDrive$)
Public Declare Function WFW_NetServerBrowseDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%)
Public Declare Function WFW_DetectNet% Lib "NPPROW16.DLL" ()
Public Declare Function WFW_NetShareAsDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%)
Public Declare Function WFW_NetStopShareDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%)
Public Declare Function WFW_I_ConnectDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal nType As Integer)
Public Declare Function WFW_I_ConnectionDialog% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal nType As Integer)
Public Declare Function WFW_I_ChangePassword% Lib "NPPROW16.DLL" (ByVal nhwnd%)
Public Declare Function WFW_I_Logoff% Lib "NPPROW16.DLL" (ByVal nhwnd%)
Public Declare Function WFW_NetGetConnection% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal cDrive$, ByVal cDevice$)
Public Declare Function WFW_NetGetShareCount% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal nType As Integer)
Public Declare Function WFW_I_ChangeCachePassword% Lib "NPPROW16.DLL" (ByVal nhwnd%)
Public Declare Function WFW_NetGetUser% Lib "NPPROW16.DLL" (ByVal nhwnd%, ByVal cUName$)
Public Declare Function WFW_NetDDEGetNodeName% Lib "NPPROW16.DLL" (ByVal cUName$)
Public Declare Function WFW_NetGetNetworkProfile% Lib "NPPROW16.DLL" (ByVal cUName$, ByVal cCName$, ByVal cWrkGrp$, ByVal cComment$)


