/*
Copyright(C) Delcom-Deltranik International Software Engineering 1990-1993.

$owner: BILLW$

$version: 1.14$ $date: February 24, 1993$ $time: 03:20:09 PM$

$lgb$
1.0 October 22, 1992 BILLW
1.1 October 22, 1992 BILLW
1.2 October 23, 1992 BILLW
1.3 October 23, 1992 BILLW
1.4 October 30, 1992 BILLW
1.5 November 13, 1992 BILLW
1.6 December 3, 1992 BILLW
1.7 December 3, 1992 BILLW
1.8 December 3, 1992 BILLW
1.9 December 3, 1992 BILLW
1.10 December 23, 1992 BILLW
1.11 December 23, 1992 BILLW
1.12 February 12, 1993 BILLW
1.13 February 19, 1993 BILLW
1.14 February 24, 1993 BILLW
$lge$
*/

#DEFINE VERSION "Version: $release: relsub1$"

/* $nokeywords$ */

#DEFINE R_RELEASE	"Release: 1.0"

#DEFINE GE_VALID     	9     // flush and exit with validation
#DEFINE GE_RIGHT     	10
#DEFINE GE_LEFT      	11

#DEFINE B_NCURRENT    1
#DEFINE B_NLISTSIZE   2
#DEFINE B_LHORIZONTAL 3
#DEFINE B_LHOTKEYS    4
#DEFINE B_LINDEXED    5
#DEFINE B_NNEWROWPOS  6
#DEFINE B_NOLDROW     7
#DEFINE B_NOLDCOL     8
#DEFINE B_NOLDCURRENT 9
#DEFINE B_NTOP        10
#DEFINE B_NLEFT       11
#DEFINE B_NBOTTOM     12
#DEFINE B_NRIGHT      13
#DEFINE B_XSEEKVAL    14
#DEFINE B_BEMPTY      15
#DEFINE B_LMORE       16
#DEFINE B_BCONDITION  17
#DEFINE B_LALLLINE    18
#DEFINE B_AKEYS       19
#DEFINE B_LMENU       20
#DEFINE B_PROMPTLIST  21
#DEFINE B_PROMPTPOS   22
#DEFINE B_ACBLOCKS    23
#DEFINE B_BCUSTOMHIGH 24
#DEFINE B_XBOTTOMVAL  25
#DEFINE B_BTHEBROWSE  26
#DEFINE B_LINDATA     27
#DEFINE B_AARRAY      28
#DEFINE B_NBTYPE      29
#DEFINE B_BEVENT      30
#DEFINE B_LVERTMOVE   31

#DEFINE GR_NAME 			1
#DEFINE GR_MESSAGE		2
#DEFINE GR_ORIGINAL		3
#DEFINE GR_COLORS			4
#DEFINE GR_WITH			5
#DEFINE GR_HIDDEN			6
#DEFINE GR_ORIGIN			7
#DEFINE GR_AREA         8
#DEFINE GR_LIST			9
#DEFINE GR_EXPAND			10

#DEFINE LB_HASFOCUS    10
#DEFINE LB_MAXROW      11
#DEFINE LB_MAXCOL      12
#DEFINE LB_BROWSE      13
#DEFINE LB_COLORS      14
#DEFINE LB_ELEMENT     15
#DEFINE LB_POSITION    16

#DEFINE FALSE				.F.
#DEFINE TRUE				.T.

#DEFINE C_AVAILABLE		1
#DEFINE C_INFOCUS			2
#DEFINE C_UNAVAILABLE	3

#DEFINE L_FILENAME		1
#DEFINE L_INDEX			2
#DEFINE L_SEEK_EXP		3
#DEFINE L_RETURN_EXP		4
#DEFINE L_DISPLAY_EXP	5
#DEFINE L_DATA_EXP   	6

#DEFINE R_BUTTON 			CHR(9)
#DEFINE R_CHECK 			CHR(251)

#DEFINE R_VALUE			1
#DEFINE R_PROMPT			2
#DEFINE R_ROW				3
#DEFINE R_COL				4


#DEFINE K_UNDO          K_CTRL_U

#DEFINE M_LEFT_BUTTON	-48
#DEFINE M_RIGHT_BUTTON	-49
#DEFINE M_MIDDLE_BUTTON	-50
#DEFINE M_DOUBLE_CLICK	-51

#COMMAND QUIT => SetMouse(0) ; ClearMouse() ; 
                 ; SETMODE( MAXROW() + 1, MAXCOL() + 1 ) ; __Quit()

#COMMAND DEFAULT <var> TO <value> => <var> := IIF(<var> == NIL, <value>, <var>)
#COMMAND DEFAULT <var> TO <value> FROM <default> ;
         => <var> := IIF(<default> == NIL, <value>, <default> )

***
*  READ
*

#COMMAND READ <nExp>                                                    ;
      => ReadModal(GetList,<nExp>)                                      ;
       ; GetList := {}

#COMMAND READ SAVE                                                      ;
      => ReadModal(GetList)

#COMMAND READ SAVE [<nExp>]                                                     ;
      => ReadModal(GetList,[<nExp>])

#COMMAND @ <row>, <col> GET <var>                                       ;
                        [USING <reader>]                                ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [<hide: HIDDEN>]                                ;
                        [COLOR <color>]                                 ;
                        [MASTERVAR <origin>]                            ;
                        [MASTERAREA <dbf>]                              ;
                        [LIST <list>]                                   ;
                        [DBLIST <dblist>]                               ;
                        [USERLIST <userlist>]                           ;
                        [PICTURE <picture>]                             ;
                        [MESSAGE <saymsg>]                              ;
                        [WITH <with>]                                   ;
                        [SEND <msg>]                                    ;
                                                                        ;
      => AAdd(                                                          ;
               GetList,                                                 ;
               __GET_( <row>, <col>,                                    ;
                      { |x| iif( x == NIL, <var>, <var> := x ) },       ;
                      <(var)>, <"reader">, <{valid}>, <{when}>,         ;
                      <.hide.>, <color>, <{origin}>, <"dbf">, <list>,   ;
                      <dblist>, <userlist>, <picture>, <saymsg>, <with> ;
							)                                                  ;
             )                                                          ;
      [; ATail(GetList):<msg>]                                          ;
      [; ATail(GetList):Reader := { |get,x,y| <reader>(get,x,y) }]      ;
      [; <reader>(ATail(GetList),.T.)]


#COMMAND @ <row>, <col> BUTTON <button>                                 ;
                        [MESSAGE <saymsg>]                              ;
                        [COLOR <color>]                                 ;
                        [ACTION <act>]                                  ;
                        [WHEN <when>]                                   ;
                        [<hide: HIDDEN>]                                ;
                                                                        ;
      => AADD( GetList,                                                 ;
               __GET_( <row>, <col>, <{button}>, <button>, "BUTTON",    ;
                     <{act}>, <{when}>, <.hide.>, <color>,,,,,,,<saymsg>);
             )                                                          ;
		   ; ATail(GetList):Reader := { |get,x| Button(get,x) }           ;
		   ; Button(ATail(GetList),.T.)


#COMMAND @ <row>, <col> LISTBOX <var>                                   ;
								TO <maxrow>,<maxcol>                            ;
								WITH <list>                                     ;
								[MESSAGE <saymsg>]                              ;
								[MASTERVAR <origin>]                            ;
								[MASTERAREA <dbf>]                              ;
								[COLOR <color>]                                 ;
								[ACTION <act>]                                  ;
								[WHEN <when>]                                   ;
																								;
		=> AADD( GetList,                                                 ;
					__GET_( <row>, <col>,                                    ;
							{ |x| iif( x == NIL, <var>, <var> := x ) },        ;
							<(var)>, "LISTBOX", <{act}>, <{when}>, .F.,        ;
							,,<{origin}>, <"dbf">,,,,<saymsg>, <list> )        ;
				 )                                                          ;
			; ATail(GetList):reader := { |get,x| ListBox( get, x ) }       ;
			; AADD( ATAIL( GetList ):cargo, .F. )                          ;
			; AADD( ATAIL( GetList ):cargo, <maxrow> )                     ;
			; AADD( ATAIL( GetList ):cargo, <maxcol> )                     ;
			; AADD( ATAIL( GetList ):cargo, NIL )                          ;
			; AADD( ATAIL( GetList ):cargo, <color> )                      ;
			; AADD( ATAIL( GetList ):cargo, 1 )                            ;
			; AADD( ATAIL( GetList ):cargo, 1 )                            ;
			; ListBox( ATail(GetList), .T. )


***
*  MENU TO / PROMPT
*

#COMMAND MENU TO <v> USING <x> [WITH <w>] [<force: FORCE>]              ;
      => <v> := PromptModal( PromptList, <x>, <w>, <.force.> )          ;
		   ; PromptList := {}

#COMMAND MENU TO <v> SAVE USING <x> [WITH <w>] [<force: FORCE>]         ;
      => <v> := PromptModal( PromptList, <x>, <w>, <.force.> )

#COMMAND MENU TO <v> [WITH <w>] [<force: FORCE>]                        ;
      => <v> := PromptModal( PromptList, <v>, <w>, <.force.> )          ;
		   ; PromptList := {}

#COMMAND MENU TO <v> SAVE [WITH <w>] [<force: FORCE>]                   ;
      => <v> := PromptModal( PromptList, <v>, <w>, <.force.> )


#COMMAND @ <row>, <col> PROMPT <prompt>                                 ;
                        [MESSAGE <msg>]                                 ;
                        [ACTION <act>]                                  ;
                        [PULLDOWN <pd>]                                 ;
                        [WHEN <wen>]                                    ;
                                                                        ;
      => AADD( PromptList, GetNew( <row>, <col>, , , ,))                ;
             ;Atail(PromptList):cargo := {"",0,0,""}                    ;
             [;Atail(PromptList):cargo := {<msg>,0,0,""}]               ;
             ;_SetPrompt( Atail(PromptList), <prompt> )                 ;
             ;Atail(PromptList):colorSpec := SetMenuColor(TRUE)         ;
             [;Atail(PromptList):postBlock := <{act}>]                  ;
             [;Atail(PromptList):postBlock := <pd>]                     ;
             [;Atail(PromptList):preBlock := <{wen}>                    ;
             ;Atail(PromptList):colorSpec:=SetMenuColor(<wen>)]         ;
             ;PromptDisp( Atail(PromptList) )

***
*  Clear screen
*

#command CLS                                                            ;
      => MScroll()                                                       ;
       ; SetPos(0,0)

#command @ <row>, <col>                                                 ;
      => MScroll( <row>, <col>, <row> )                                  ;
       ; SetPos( <row>, <col> )

#command @ <top>, <left> CLEAR                                          ;
      => MScroll( <top>, <left> )                                        ;
       ; SetPos( <top>, <left> )


#command @ <top>, <left> CLEAR TO <bottom>, <right>                     ;
      => MScroll( <top>, <left>, <bottom>, <right> )                     ;
       ; SetPos( <top>, <left> )


#command NETUSE <(db)>                                                     ;
             [VIA <rdd>]                                                ;
             [ALIAS <a>]                                                ;
             [<new: NEW>]                                               ;
             [<ex: EXCLUSIVE>]                                          ;
             [<sh: SHARED>]                                             ;
             [<ro: READONLY>]                                           ;
             [INDEX <(index1)> [, <(indexn)>]]                          ;
                                                                        ;
      => IF dbNetUse(                                                   ;
                    <.new.>, <rdd>, <(db)>, <(a)>,                      ;
                    if(<.sh.> .or. <.ex.>, !<.ex.>, NIL), <.ro.>        ;
                  )                                                     ;
      [; dbSetIndex( <(index1)> ) ]                                     ;
      [; dbSetIndex( <(indexn)> ) ]                                     ;
		 ; ELSE ; BREAK( <(db)> ) ; ENDIF




