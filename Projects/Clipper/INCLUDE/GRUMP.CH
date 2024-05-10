/*
   File: GRUMP.CH
   Excerpted version of primary Grumpfish Library header file
   Author: Greg Lief
   Copyright (c) 1988-92, Greg Lief
*/

#define CRLF          CHR(13)+CHR(10)

/* ­the ubiquitous Center() function! */
#xtranslate CENTER(<row>, <msg>, <width> [, <color>] ) => ;
            DevPos( <row>, int(( <width> - len( <msg> )) / 2)) ; ;
            DevOut( <msg> [,<color>] )

#xtranslate CENTER(<row>, <msg> [, , <color> ] ) => ;
            DevPos( <row>, int((maxcol() + 1 - len( <msg> )) / 2)) ; ;
            DevOut( <msg> [, <color>] )

/*
   a modified CENTER(), which uses setpos() and dispout() to ensure
   that output always goes to the screen and never the printer
*/
#xtranslate SCRNCENTER(<row>, <msg> [, <color> ] ) => ;
            SetPos( <row>, int(( maxcol()+1 - len(<msg>)) / 2)) ; ;
            DispOut( <msg> [, <color>] )

/*
   a modified @..SAY which uses setpos() and dispout() to ensure
   that output always goes to the screen and never the printer
*/
#xcommand @ <row>, <col> SSAY <xpr> [COLOR <color>] => ;
         SetPos( <row>, <col> ) ; DispOut( <xpr> [, <color>] )

/* constants for COLORSET() */
#define C_GETS                       1
#define C_MESSAGE                    2
#define C_MESSAGEBLINK               3
#define C_WAITMESSAGE                4
#define C_WAITMESSAGEBLINK           5
#define C_YESNO                      6
#define C_ERRORBOX                   7
#define C_ERRORMESSAGE               8
#define C_MEMEDIT_BOX                9
#define C_MEMEDIT_WINDOW            10
#define C_STOPWATCH_BOX             11
#define C_STOPWATCH_WINDOW          12
#define C_CALCULATOR_BOX            13
#define C_CALCULATOR_WINDOW         14
#define C_NOTEPAD_BOX               15
#define C_NOTEPAD_WINDOW            16
#define C_PHONEBOOK_WINDOW1         17
#define C_PHONEBOOK_WINDOW2         18
#define C_CALENDAR                  19
#define C_APPOINTMENT               20
#define C_APICK_BOXOUTLINE          21
#define C_APICK_STATUSBAR           22
#define C_APICK_INDICATOR           23
#define C_APICK_UNSELECTED          24
#define C_APICK_CURRENT             25
#define C_APICK_TAGGED              26
#define C_APICK_TAGGEDCURRENT       27
#define C_APICK_UNAVAILABLE         28
#define C_APICK_UNAVAILABLECURRENT  29
#define C_MENU_UNSELECTED           30
#define C_MENU_SELECTED             31
#define C_GRUMPBROW_BOX             32
#define C_GRUMPBROW_SAY             33
#define C_GRUMPBROW_GET             34

#xtranslate AKill( <a>, <e> ) => asize( adel( <a>, <e> ), len( <a> ) - 1 )

#xtranslate Truncate( <a> ) => asize( <a>, len(<a>) - 1)

#xcommand waitoff( [<sound:ON,NOISE,SOUND>] ) => WaitOn(.F., <.sound.>)

#xtranslate StripExt( <fname> ) => ;
      if('.' $ <fname>, substr( <fname>, 1, at('.', <fname>) - 1), <fname> )

#xcommand DEFAULT <param> TO <value> [, <param2> TO <value2> ] => ;
          IF <param> == NIL ; <param> := <value> ; END ;
          [; IF <param2> == NIL ; <param2> := <value2> ; END ]

#xtranslate ByeByeBox(<buff>) => ;
            restscreen(<buff>\[1], <buff>\[2], <buff>\[3], <buff>\[4], ;
                       <buff>\[5])

* eof GRUMP.CH
