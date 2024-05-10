/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
³                     Ganahl Software, Inc.                      ³
³                        (305) 484-8394                          ³
³                                                                ³
³          (c) 1990-1994 Copyright Ganahl Software, Inc.         ³
³                      All Rights Reserved                       ³
³                                                                ³
ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
#ifndef FLEXERR_CH_
   #define FLEXERR_CH_

   #define E_NoError                    0
   #define E_NoOpenDBV                  1
   #define E_BadDBVversion              2
   #define E_BadType                    3
   #define E_BadTypeInDBV               4
   #define E_BadAlias                   5
   #define E_OutOfMemory                6
   #define E_ListOverflow               7
   #define E_ArrayRecursive             8
   #define E_NoOpenDBF                  9
   #define E_StringOver65k             10
   #define E_CompressFail              11
   #define E_NotValidCmpStr            12
   #define E_ExpandFail                13
   #define E_NotFFCompressed           14
   #define E_CreateFile                15
   #define E_CommitCache               16
   #define E_ArrayWriteLen             17
   #define E_HdrLock                   18
   #define E_HdrRead                   19
   #define E_UseLock                   20
   #define E_NotAFlexFile              21
   #define E_TopLock                   22
   #define E_BadArea                   25
   #define E_AreaOutOfBounds           26
   #define E_FilesOpen                 27
   #define E_ChkDeletedFailed          28
   #define E_FreeSpaceFailed           29
   #define E_GetCacheFailed            30
   #define E_TokenSignature            31
   #define E_TokenType                 32
   #define E_TokenDeleted              33
   #define E_TooFewParams              34
   #define E_InvalidParam              35

   #define E_BadDBFToken               41
   #define E_IllegalTokenType          42
   #define E_InvalidClassName          43


   #define E_TempNameFail              46
   #define E_FileNameTooLong           47
   #define E_ArrayElementBad           48

   #define E_WriteOutOfBounds          50
   #define E_ReadOutOfBounds           51
   #define E_INV_REC_SZ                52
   #define E_INV_DBF_HDR               53

   #define E_MissingExternal           65
   #define E_ReturnValue               66
   #define E_MemoryCorrupt             67
   #define E_ArraySize                 68
   #define E_ParmTypeError             69
   #define E_CacheFail                 70
   #define E_NoHuffTbl                 71
   #define E_DuplicateAlias            72
   #define E_ArrayRetrieve             73
   #define E_87ArrayDate               74
   #define E_ArrayCharTooLong          75
   #define E_ArrayTooLarge             76
   #define E_NoFlexDriver              77

   #define E_FileAccess                80
   #define E_ReadError                 81
   #define E_CreateError               82
   #define E_OpenError                 83
   #define E_WriteError                84
   #define E_NothingMoreToSay          85

#endif
