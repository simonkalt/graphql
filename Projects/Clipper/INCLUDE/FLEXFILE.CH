/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
³                     Ganahl Software, Inc.                      ³
³                        (305) 484-8394                          ³
³                                                                ³
³          (c) 1990-1994 Copyright Ganahl Software, Inc.         ³
³                      All Rights Reserved                       ³
³                                                                ³
ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
³   The header file for FlexFile II is purposely small.  Because ³
³   of the inherent limitation of the CA-Clipper pre-processor   ³
³   becoming overbooked, we have limitted the pre-processor      ³
³   directives to the following:                                 ³
ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
#ifndef FLEXFILE_
   #define FLEXFILE_

   #define USE_EXCLUSIVE      .t. // fourth parameter of v_use()
   #define USE_SHARED         .f. // fourth parameter of v_use()
   #define NEW_AREA           .t. // Third parameter of v_use()
   #define ERR_OFF            ""

   // DBV file open mode (use v_or() to OR these together) or use v_openmode()
   #define V_READONLY         00
   #define V_READWRITE        02
   #define V_EXCLUSIVE        16
   #define V_SHARED           64

   #define CMP_DEFAULT         0  // Use V_SETCMP() setting.
   #define CMP_NONE            1  // Turn Automatic compression off (default)
   #define CMP_SPEED           2  // Very fast (Effective on English text only)
   #define CMP_SIZE            3  // Maximum compression (FlexFile chooses)
   #define CMP_SCREEN          4  // Very fast screen compression
   #define CMP_LZ              5  // Ziv Limpel compression (Pkzip-ish)
   #define CMP_CUSTOM          6  // Customized Fast compression (see MakeHuff)

   #define V2F_APPEND         'A' // Append data to target file
   #define V2F_OVERWRITE      'O' // Overwrite target file
   #define V2F_BINARY         'B' // Target is LPT, set in binary mode

   #define V_EXP2TARGET        0
   #define V_CPY2TARGET        1
   #define V_CMP2TARGET        2

   #define QRY_LITERAL        .t. // Search for <cFindThis> literally
   #define QRY_INTERPRET      .f. // Interpret <cFindThis> for wildcards and logic
   #define QRY_CASE_SENSITIVE .t. // Case sensitive search regardless of <lMode>
   #define QRY_NO_CASE        .f. // Case insensitive
   #define QRY_WILD_CARDS     .t. // Slightly slower than NO_WILD_CARDS
   #define QRY_NO_WILD_CARDS  .f. // Slightly faster than WILD_CARDS
   #define QRY_FIRST          .t.
   #define QRY_NEXT           .f.

#endif
