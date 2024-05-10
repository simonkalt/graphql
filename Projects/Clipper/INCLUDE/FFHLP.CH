/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
³                     Ganahl Software, Inc.                      ³
³                        (305) 484-8394                          ³
³                                                                ³
³          (c) 1990-1994 Copyright Ganahl Software, Inc.         ³
³                      All Rights Reserved                       ³
³                                                                ³
ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
// for developer to access the Help Settings Array
// which is stored in v_TopPtr()

#define VH_MODE_SHORT                       0
#define VH_MODE_LONG                        1

// SETTINGS structure
// not editable at all!
#define VH_SET_VERSION                      001  // not editable

// not manually editable
#define VH_SET_SHORT_TOP                    005  // not editable
#define VH_SET_SHORT_LEFT                   006  // not editable
#define VH_SET_SHORT_BOTTOM                 007  // not editable
#define VH_SET_SHORT_RIGHT                  008  // not editable

// editable
#define VH_SET_FRAME_MOVE                   010  //
#define VH_SET_FRAME_MOVE_COLOR             011  //
#define VH_SET_FRAME_SIZE                   012  //
#define VH_SET_FRAME_SIZE_COLOR             013  //

#define VH_SET_TONE_FREQ                    014  //
#define VH_SET_TONE_LEN                     015  //

#define VH_SET_SHADOW                       016  //
#define VH_SET_SHADOW_COLOR                 017  //
#define VH_SET_COLOR                        018  //
#define VH_SET_SOUND                        019  //

#define VH_SET_CURSOR_MODE_INS              020  //
#define VH_SET_CURSOR_INS                   021  //
#define VH_SET_CURSOR_OVR                   022  //

#define VH_SET_AUTOCLOSE                    023  //
#define VH_SET_SCROLLBAR_CHAR               024  //
#define VH_SET_SCROLLBUTTON_CHAR            025  //

#define VH_SET_MSG_FRAME                    026  //
#define VH_SET_MSG_FRAME_COLOR              027  //
#define VH_SET_MSG_COLOR                    028  //
#define VH_SET_MSG_TEXT_COLOR               029  //

#define VH_SET_HELP_FRAME_COLOR             030  //
#define VH_SET_HELP_FRAME                   031  //
#define VH_SET_HELP_COLOR                   032  //
#define VH_SET_HELP_TEXT_COLOR              033  //
#define VH_SET_HELP_SEL_COLOR               034  //
#define VH_SET_HELP_UNSEL_COLOR             035  //
#define VH_SET_HELP_SCROLLBAR_COLOR         036  //
#define VH_SET_HELP_SCROLLBUTTON_COLOR      037  //

// Settings specific to DEV obj
#define VH_SET_DEV_SHORT_FRAME_COLOR        050  //
#define VH_SET_DEV_SHORT_FRAME              051  //
#define VH_SET_DEV_SHORT_COLOR              052  //
#define VH_SET_DEV_SHORT_TEXT_COLOR         053  //
#define VH_SET_DEV_SHORT_SEL_COLOR          054  //
#define VH_SET_DEV_SHORT_UNSEL_COLOR        055  //

#define VH_SET_DEV_LONG_FRAME               060  //
#define VH_SET_DEV_LONG_FRAME_COLOR         061  //
#define VH_SET_DEV_LONG_COLOR               062  //
#define VH_SET_DEV_LONG_TEXT_COLOR          063  //
#define VH_SET_DEV_LONG_EDITOR_COLOR        064  //
#define VH_SET_DEV_LONG_BLOCK_COLOR         065  //
#define VH_SET_DEV_LONG_EOL_CHAR            068  //
#define VH_SET_DEV_LONG_EOF_CHAR            069  //
#define VH_SET_DEV_LONG_SCROLLBAR_COLOR     070  //
#define VH_SET_DEV_LONG_SCROLLBUTTON_COLOR  071  //

#define VH_SET_DEV_IDPICK_FRAME             075  //
#define VH_SET_DEV_IDPICK_FRAME_COLOR       076  //
#define VH_SET_DEV_IDPICK_COLOR             077  //
#define VH_SET_DEV_IDPICK_TEXT_COLOR        078  //
#define VH_SET_DEV_IDPICK_SEL_COLOR         079  //
#define VH_SET_DEV_IDPICK_UNSEL_COLOR       080  //

#define VH_SET_DEV_NEWHELP_FRAME            085  //
#define VH_SET_DEV_NEWHELP_FRAME_COLOR      086  //
#define VH_SET_DEV_NEWHELP_COLOR            087  //
#define VH_SET_DEV_NEWHELP_TEXT_COLOR       088  //
#define VH_SET_DEV_NEWHELP_SEL_COLOR        089  //
#define VH_SET_DEV_NEWHELP_UNSEL_COLOR      090  //

#define VH_SET_DEV_SETTINGS_FRAME           095  //
#define VH_SET_DEV_SETTINGS_FRAME_COLOR     096  //
#define VH_SET_DEV_SETTINGS_COLOR           097  //
#define VH_SET_DEV_SETTINGS_TEXT_COLOR      098  //
#define VH_SET_DEV_SETTINGS_SEL_COLOR       099  //
#define VH_SET_DEV_SETTINGS_UNSEL_COLOR     100  //

#define VH_SET_DEV_ATTRPICK_FRAME           101  // Not in edit settings yet
#define VH_SET_DEV_ATTRPICK_FRAME_COLOR     102  // Not in edit settings yet

#define VH_SET_DEV_FRAMEPICK_FRAME          105  // Not in edit settings yet
#define VH_SET_DEV_FRAMEPICK_FRAME_COLOR    106  // Not in edit settings yet
#define VH_SET_DEV_FRAMEPICK_SEL_COLOR      107  // Not in edit settings yet
#define VH_SET_DEV_FRAMEPICK_UNSEL_COLOR    108  // Not in edit settings yet

#define VH_SET_DEV_ASCPICK_FRAME            110  // Not in edit settings yet
#define VH_SET_DEV_ASCPICK_FRAME_COLOR      111  // Not in edit settings yet
#define VH_SET_DEV_ASCPICK_SEL_COLOR        112  // Not in edit settings yet
#define VH_SET_DEV_ASCPICK_UNSEL_COLOR      113  // Not in edit settings yet

// Settings specific to VIEW obj
#define VH_SET_USER_SHORT_FRAME             130  //
#define VH_SET_USER_SHORT_FRAME_COLOR       131  //
#define VH_SET_USER_SHORT_COLOR             132  //
#define VH_SET_USER_SHORT_TEXT_COLOR        133  //
#define VH_SET_USER_SHORT_SEL_COLOR         134  //
#define VH_SET_USER_SHORT_UNSEL_COLOR       135  //

#define VH_SET_USER_LONG_FRAME              140  //
#define VH_SET_USER_LONG_FRAME_COLOR        141  //
#define VH_SET_USER_LONG_COLOR              142  //
#define VH_SET_USER_LONG_TEXT_COLOR         143  //
#define VH_SET_USER_LONG_VIEW_COLOR         144  //
#define VH_SET_USER_LONG_SEL_COLOR          145  //
#define VH_SET_USER_LONG_UNSEL_COLOR        146  //
#define VH_SET_USER_LONG_CURSOR_COLOR       147  //
#define VH_SET_USER_LONG_SCROLLBAR_COLOR    148  //
#define VH_SET_USER_LONG_SCROLLBUTTON_COLOR 149  //

#define VH_SET_USER_HISTORY_FRAME           160  //
#define VH_SET_USER_HISTORY_FRAME_COLOR     162  //
#define VH_SET_USER_HISTORY_COLOR           163  //
#define VH_SET_USER_HISTORY_TEXT_COLOR      164  //
#define VH_SET_USER_HISTORY_SEL_COLOR       165  //
#define VH_SET_USER_HISTORY_UNSEL_COLOR     166  //

#define VH_SET_ELEMENTS                     180

// accumulate HELP ID's from all gets in getList and push onto stack
// with codeblock to tell how to get the code from each oGet
// the default block will be {|oGet| oGet:cargo}
// this should be follwed by 'the' READ for the gets in getList
// don't forget to VHPOP or vh_pop() after the READ!!!
#xcommand VHPUSHGETS [IDBLOCK <bID>] ;
       => vh_pushGets( getList, <bID> )

// this UDC combines VHPUSHGETS, READ and VHPOP
// just load the get objects with their help ID's and run this
// the push/read/pop are all done handled with this UDC
#xcommand VHREAD [IDBLOCK <bID>] ;
     =>   VHPUSHGETS [IDBLOCK <bID>] ;
        ; READ ;
        ; VHPOP

// accumulate HELP ID's from all browser columns in browser and push onto stack
// with codeblock to tell how to get the code from each oColumn
// the default block will be {|oCol| oCol:cargo}
// don't forget to VHPOP or vh_pop() after the browser exit!!!
#xcommand VHPUSHTB <browser> [IDBLOCK <bID>] ;
       => vh_push( { {|| vh_tbPos( <browser> )}, vh_LoadTB( <browser>, <bID> ) } )

// for people who don't like udf's
#xcommand VHPUSH [HELP] [ <id> ] => vh_push( <id> )
#xcommand VHPOP  [HELP]          => vh_pop()

// direct call to help via a UDC
#xcommand VHELP LONG  <id> => vh_help( <id>, VH_MODE_LONG)
#xcommand VHELP SHORT <id> => vh_help( <id>, VH_MODE_SHORT )
#xcommand VHELP SHORT      => vh_help(NIL, VH_MODE_SHORT)

// hook help to a hot key
#xcommand SET VHELP SHORT [TO] <nKey> => setKey( <nKey>, {|| vh_help(NIL, VH_MODE_SHORT)} )
#xcommand SET VHELP LONG  [TO] <nKey> => setKey( <nKey>, {|| vh_hot()} )

// saving and restoring of the stack via a UDC
// usually used around a READ if pops are done in the valid clause
// because [Esc] causes the current VALID to NOT be run, and doesn't pop back!
#xcommand VHELP STACK SAVE    => vh_stackSave()
#xcommand VHELP STACK REST    => vh_stackRest()
#xcommand VHELP STACK RESTORE => VHELP STACK REST

