/*
 * Source     : AXNTXCDX.CH for Clipper 5.01a
 * Description: Header file for both Advantage RDDs (DBFNTXAX and DBFCDXAX).
 *              Include this file in your application if you want to use
 *              both the Advantage NTX RDD (DBFNTXAX) and the Advantage
 *              CDX/IDX RDD (DBFCDXAX) in the same application and also if
 *              you want to make use of the advanced functionality that is
 *              available with the Advantage RDDs, i.e. Conditional Indexes,
 *              Compound Indexes (Tags), Sub-Indexes, SCOPEs, AUTOUSE, etc.
 *
 * Copyright 1993 - Extended Systems, Inc.
 */

EXTERN _VDBFAXS
EXTERN _VDBFCAX

#command SET INDEX TO [ <(index1)> [, <(indexn)>]]                         ;
      => dbClearIndex()                                                    ;
      [; dbSetIndex( <(index1)> )]                                         ;
      [; dbSetIndex( <(indexn)> )]

#command SET INDEX TO [ <(index1)> [, <(indexn)>]] ADDITIVE                ;
      => [dbSetIndex( <(index1)> )]                                        ;
         [; dbSetIndex( <(indexn)> )]

#command SET TAGORDER TO <order>                                           ;
      => AX_SetTagOrder( <order> )

#command SET TAGORDER TO                                                   ;
      => AX_SetTagOrder( 0 )

#command SET ORDER TO TAG <(tag)>                                          ;
         [OF <(cdx)>]                                                      ;
         [IN <(cdx)>]                                                      ;
      => AX_SetTag( <(tag)>, <(cdx)> )

#command SET TAG TO <(tag)>                                                ;
         [OF <(cdx)>]                                                      ;
         [IN <(cdx)>]                                                      ;
      => AX_SetTag( <(tag)>, <(cdx)> )

#command SET TAG TO                                                        ;
      => dbSetOrder(0)

#command INDEX ON <key> TO <(file)>                                        ;
         [FOR       <for>]                                                 ;
         [<all:     ALL>]                                                  ;
         [WHILE     <while>]                                               ;
         [NEXT      <next>]                                                ;
         [RECORD    <rec>]                                                 ;
         [<rest:    REST>]                                                 ;
         [EVAL      <eval> [EVERY  <every>]]                               ;
         [OPTION    <eval> [STEP   <every>]]                               ;
         [<unique:  UNIQUE>]                                               ;
         [<ascend:  ASCENDING>]                                            ;
         [<descend: DESCENDING>]                                           ;
         [<cur:     USECURRENT>]                                           ;
         [<cur:     SUBINDEX>]                                             ;
         [<add:     ADDITIVE>]                                             ;
         [<non:     NONCOMPACT>]                                           ;
      => _axCondSet( <"for">, <{for}>,                                     ;
                      if( <.all.>, .t., NIL ),                             ;
                      <{while}>,                                           ;
                      <{eval}>, <every>,                                   ;
                      RECNO(), <next>, <rec>,                              ;
                      if( <.rest.>, .t., NIL ),                            ;
                      if( (<.descend.> .AND. !<.ascend.>), .t., NIL ),     ;
                      .f., NIL, <.cur.>, NIL, <.non.>, <.add.>, NIL)       ;
       ; dbCreateIndex(<(file)>, <"key">, <{key}>,                         ;
                       if( <.unique.>, .t., NIL))

#command INDEX ON <key> TAG <(tag)>                                        ;
         [OF <(cdx)>]                                                      ;
         [TO <(cdx)>]                                                      ;
         [FOR        <for>]                                                ;
         [<all:      ALL>]                                                 ;
         [WHILE      <while>]                                              ;
         [NEXT       <next>]                                               ;
         [RECORD     <rec>]                                                ;
         [<rest:     REST>]                                                ;
         [EVAL       <eval> [EVERY  <every>]]                              ;
         [OPTION     <eval> [STEP   <every>]]                              ;
         [<unique:   UNIQUE>]                                              ;
         [<ascend:   ASCENDING>]                                           ;
         [<descend:  DESCENDING>]                                          ;
         [<cur:      USECURRENT>]                                          ;
         [<cur:      SUBINDEX>]                                            ;
         [<add:      ADDITIVE>]                                            ;
      => _axCondSet( <"for">, <{for}>,                                     ;
                      if( <.all.>, .t., NIL ),                             ;
                      <{while}>,                                           ;
                      <{eval}>, <every>,                                   ;
                      RECNO(), <next>, <rec>,                              ;
                      if( <.rest.>, .t., NIL ),                            ;
                      if( (<.descend.> .AND. !<.ascend.>), .t., NIL ),     ;
                      .t., <(cdx)>, <.cur.>, NIL, .t., <.add.>, NIL)       ;
       ; dbCreateIndex(<(tag)>, <"key">, <{key}>,                          ;
                       if( <.unique.>, .t., NIL))


#command SUBINDEX ON <key> TO <(file)>                                     ;
         [FOR       <for>]                                                 ;
         [<all:     ALL>]                                                  ;
         [WHILE     <while>]                                               ;
         [NEXT      <next>]                                                ;
         [RECORD    <rec>]                                                 ;
         [<rest:    REST>]                                                 ;
         [EVAL      <eval> [EVERY  <every>]]                               ;
         [OPTION    <eval> [STEP   <every>]]                               ;
         [<unique:  UNIQUE>]                                               ;
         [<ascend:  ASCENDING>]                                            ;
         [<descend: DESCENDING>]                                           ;
         [<non:     NONCOMPACT>]                                           ;
         [<add:     ADDITIVE>]                                             ;
                                                                           ;
      => _axCondSet(<"for">, <{for}>,                                      ;
                      if( <.all.>, .t., NIL ),                             ;
                      <{while}>,                                           ;
                      <{eval}>, <every>,                                   ;
                      RECNO(), <next>, <rec>,                              ;
                      if( <.rest.>, .t., NIL ),                            ;
                      if( (<.descend.> .AND. !<.ascend.>), .t., NIL ),     ;
                      .f., NIL, .t., NIL, <.non.>, <.add.>, NIL)           ;
       ; dbCreateIndex(<(file)>, <"key">, <{key}>,                         ;
                        if(<.unique.>, .t., NIL))

#command SUBINDEX ON <key> TAG <(tag)>                                     ;
         [OF <(cdx)>]                                                      ;
         [TO <(cdx)>]                                                      ;
         [FOR        <for>]                                                ;
         [<all:      ALL>]                                                 ;
         [WHILE      <while>]                                              ;
         [NEXT       <next>]                                               ;
         [RECORD     <rec>]                                                ;
         [<rest:     REST>]                                                ;
         [EVAL       <eval> [EVERY  <every>]]                              ;
         [OPTION     <eval> [STEP   <every>]]                              ;
         [<ascend:   ASCENDING>]                                           ;
         [<descend:  DESCENDING>]                                          ;
         [<unique:   UNIQUE>]                                              ;
         [<add:      ADDITIVE>]                                            ;
                                                                           ;
      => _axCondSet(<"for">, <{for}>,                                      ;
                      if( <.all.>, .t., NIL ),                             ;
                      <{while}>,                                           ;
                      <{eval}>, <every>,                                   ;
                      RECNO(), <next>, <rec>,                              ;
                      if( <.rest.>, .t., NIL ),                            ;
                      if( (<.descend.> .AND. !<.ascend.>), .t., NIL ),     ;
                      .t., <(cdx)>, .t., NIL, .t., <.add.>, NIL)           ;
       ; dbCreateIndex(<(tag)>, <"key">, <{key}>,                          ;
                        if(<.unique.>, .t., NIL))

#command REINDEX                                                           ;
         [EVAL       <eval> [EVERY  <every>]]                              ;
         [OPTION     <eval> [STEP   <every>]]                              ;
      => _axCondSet(,,,, <{eval}>, <every>,,,,,,,,,,,)                     ;
       ; dbReindex()

#command DELETE TAG <(tag1)>                                               ;
         [OF <(cdx1)>]                                                     ;
         [, <(tagn)> [OF <(cdxn)>]]                                        ;
      => AX_KillTag( <(tag1)>, <(cdx1)> )                                  ;
      [; AX_KillTag( <(tagn)>, <(cdxn)> )]

#command DELETE TAG <(tag1)>                                               ;
         [IN <(cdx1)>]                                                     ;
         [, <(tagn)> [IN <(cdxn)>]]                                        ;
      => AX_KillTag( <(tag1)>, <(cdx1)> )                                  ;
      [; AX_KillTag( <(tagn)>, <(cdxn)> )]

#command DELETE TAG ALL                                                    ;
         [OF <(cdx)>]                                                      ;
         [IN <(cdx)>]                                                      ;
      => AX_KillTag( .t., <(cdx)> )

// Clear both values
#command CLEAR SCOPE                                                       ;
      => AX_ClrScope(0)                                                    ;
       ; AX_ClrScope(1)

// First value, inclusive
#xcommand SET SCOPETOP TO <value>                                          ;
      => AX_SetScope(0, <value>)

#xcommand SET SCOPETOP TO                                                  ;
      => AX_ClrScope(0)

// Last value, inclusive
#xcommand SET SCOPEBOTTOM TO <value>                                       ;
      => AX_SetScope(1, <value>)

#xcommand SET SCOPEBOTTOM TO                                               ;
      => AX_ClrScope(1)

// Clear both values
#command SET SCOPE TO                                                      ;
      => AX_ClrScope(0)                                                    ;
       ; AX_ClrScope(1)

// Both values, inclusive
#command SET SCOPE TO <value>                                              ;
      => AX_SetScope(0, <value>)                                           ;
       ; AX_SetScope(1, <value>)

#command SET MEMOBLOCK TO <value>                                          ;
      => AX_SetMemoBlock(<value>)

#command SORT [TO <(file)>] [ON <fields,...>]                              ;
         [FOR    <for>]                                                    ;
         [WHILE  <while>]                                                  ;
         [NEXT   <next>]                                                   ;
         [RECORD <rec>]                                                    ;
         [<rest: REST>]                                                    ;
         [ALL]                                                             ;
         [<cur:  USECURRENT>]                                              ;
                                                                           ;
      => AX_SortOption(<.cur.>)                                            ;
       ; __dbSort(<(file)>, { <(fields)> },                                ;
                   <{for}>, <{while}>, <next>, <rec>, <.rest.>)

#command AUTOUSE <(db)> VIA <rdd> ALTERNATE <altrdd>                       ;
         [ALIAS <a>]                                                       ;
         [<new: NEW>]                                                      ;
         [<ex: EXCLUSIVE>]                                                 ;
         [<sh: SHARED>]                                                    ;
         [<ro: READONLY>]                                                  ;
         [INDEX <(index1)> [, <(indexn)>]]                                 ;
      =>                                                                   ;
        IF  AX_Loaded( AxGetDrive( <(db)> ) )                              ;
      ;    dbUseArea(                                                      ;
                     <.new.>, <rdd>, <(db)>, <(a)>,                        ;
                     if (<.sh.> .or. <.ex.>, !<.ex.>, NIL), <.ro.>         ;
                    )                                                      ;
           [; dbSetIndex( <(index1)> )]                                    ;
           [; dbSetIndex( <(indexn)> )]                                    ;
      ; ELSE                                                               ;
      ;    dbUseArea(                                                      ;
                    <.new.>, <altrdd>, <(db)>, <(a)>,                      ;
                    if (<.sh.> .or. <.ex.>, !<.ex.>, NIL), <.ro.>          ;
                    )                                                      ;
           [; dbSetIndex( <(index1)> )]                                    ;
           [; dbSetIndex( <(indexn)> )]                                    ;
      ; ENDIF
