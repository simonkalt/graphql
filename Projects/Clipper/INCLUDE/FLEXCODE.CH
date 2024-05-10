/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
³                     Ganahl Software, Inc.                      ³
³                        (305) 484-8394                          ³
³                                                                ³
³          (c) 1990-1994 Copyright Ganahl Software, Inc.         ³
³                      All Rights Reserved                       ³
³                                                                ³
ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
#ifndef FLEXCODE_
   #define FLEXCODE_

   /*ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
   ³   Use these codes to indicate the calling procedure in         ³
   ³   the event of an error.  Add to these numbers the             ³
   ³   number within the routine which caused the error.            ³
   ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
   #define E__INIT_CACHE                1
   #define E__USE                       2
   #define E__COMMIT_CACHE              3
   #define EA_USE                       4
   #define E__WRITE_TOKEN               5
   #define E__DBF2TOKEN                 6
   #define ER_REPLACE_RDD               7
   #define EA_MEMOPUT                   8
   #define E__SELECT                    9
   #define E__FLUSH_BUF                10
   #define EA_REPLACE                  11
   #define EA_SETCMP                   12
   #define EA_WALK_DBV                 13
   #define EA_FILE2VLF                 14
   #define EA_CREATE                   15
   #define E__EXPFILE                  16
   #define EI_ARR_FILL                 17
   #define ES_INIT_BUFF                18
   #define ES_GET_FROM_LIST            19
   #define ES_INIT_LIST                20
   #define ES_PUSH_LIST                21
   #define E__REPLACE_STREAM           22
   #define E__INS_DATA                 23
   #define EA_SIGNATURE                24
   #define ER_LEN_RDD                  25
   #define EI_MAKEOBJECT               26
   #define EA_MEMOGET                  27
   #define EA_RETRIEVE                 28
   #define E__RETRIEVE_STREAM          29
   #define E__COPYFILE                 30
   #define E__FILES                    31
   #define EA_SAVEBOX                  32
   #define E__HDR_LOCK                 33
   #define E__BYTE_LOCK                34
   #define EI_ARRPUT_GUTS              35
   #define ES_INIT_FILE                36
   #define ES_FLUSH_BUFF               37
   #define EA_DECRYPT                  38
   #define EA_PUTMEMO                  39
   #define EA_ENCRYPT                  40
   #define E__TOPLOCK                  41
   #define E__TOPUNLOCK                42
   #define E__TOPPTR                   43
   #define EA_CMPSTR                   44
   #define E__FREE                     45
   #define ES_TESTTOKEN                46
   #define EA_PTR                      47
   #define EA_SETKEY                   48
   #define EA_EXPSTR                   49
   #define EA_EXPFILE                  50
   #define EA_CMPFILE                  51
   #define E__CMPFILE                  52
   #define E__CHK_INTEGRITY            53
   #define E__CHK_INTEGRITY_GUTS       54
   #define E__ALIAS2AREA               55
   #define E__FILE2VLF                 56
   #define E__SET_AREA                 57
   #define ER_RETRIEVE_RDD             59
   #define ER_USE                      60
   #define E__CREATE                   61
   #define E__VLFTOKEN                 62
   #define EI_REPLACE_ARRAY            63
   #define E__VLF2FILE                 64
   #define EA_QRY_DBF                  65
   #define EA_QRY_GEN                  66
   #define EA_QRY_DBV                  67
   #define EA_VLF2FILE                 68
   #define E__QUERY                    70
   #define E__POST_FIX                 71
   #define E__GET_NEXT                 72
   #define E__CHK_CACHE                73
   #define E__RECOVER                  74
   #define EA_ARR_INIT                 75
   #define EA_MEMO2VLF                 76
   #define EA_SELECT                   77
   #define EI_CACHE_LOCK               78
   #define EI_CACHE_UNLOCK             79
   #define EA_POKE                     80
   #define E__POKE                     81
   #define EA_STUFF                    82
   #define E__STUFF                    83
   #define E__ZAP                      84
   #define EA_CLOSE                    85
   #define EE_ARRPUT                   86
   #define EA_FILL_ARR                 87
   #define EA_FIXED                    88
   #define EA_RESTBOX                  89
   #define EI_ARR_INIT                 90
   #define EI_ARR_NEXT                 91
   #define EA_LEN                      92
   #define EA_FILLARR                  93
   #define EA_DELETE                   94
   #define EA_INTEGRITY                95
   #define EA_RECOVER                  96
   #define EA_TOPLOCK                  97
   #define EA_TOPPTR                   98
   #define EA_TYPE                     99
   #define EA_ZAP                     100
   #define ER_PUTVALUE                101
   #define EA_REC_LOCK                102
   #define EA_REC_UNLOCK              103
   #define EA_LOCK                    104
   #define EA_UNLOCK                  105
   #define E_DBV2DBV                  106
   #define EA_DBV2DBV                 107
   #define ER_GETVARLEN               108
   #define E__GET_CACHE               109
   #define E__PUT_CACHE               110

#endif
