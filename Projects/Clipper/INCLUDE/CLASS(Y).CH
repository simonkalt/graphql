/*
        class(y).ch - v1.03rt

        Copyright (c) 1991 Anton van Straaten.
        All rights reserved.

        Header file for the runtime version of Class(y),
        the object oriented extension to Clipper 5.01.
*/

#define CLASSY_VER100


#xcommand   CLASS <className> [FROM <superClass>]   ;
                =>                                  ;
                CREATE CLASS <className> [FROM <superClass>]

#xcommand   CREATE CLASS <className> [FROM <superClass>]    ;
                =>                                          ;
                function <className>                        ;;
                static self                                 ;;
                local clsMkr                                ;;
                if self == NIL                              ;;
                clsMkr := ClassMaker():new(<"className">, [<superClass>()]) ;;
                message new constructor                     ;;
                hidden:

#command    INITIALIZE [CLASS] :    =>  self := clsMkr:makeClass; csyMethify(self)

#xcommand   END CLASS                           ;
                =>                              ;
                if self == nil                  ;;
                    self := clsMkr:makeClass    ;;
                else                            ;;
                    csyMethify()                ;;
                end; end                        ;;
                return self                     ;;
                #include "endclass.ch"

#xcommand   ENDCLASS    =>      END CLASS

#command    EXPORTED:   =>      ClassMaker():scope(1)
#command    VISIBLE:    =>      ClassMaker():scope(1)

#command    HIDDEN:     =>      ClassMaker():scope(2)
#command    LOCAL:      =>      HIDDEN:
#command    PRIVATE:    =>      HIDDEN:

#command    PROTECTED:  =>      ClassMaker():scope(4)

#command    VAR <(name1)> [, <(nameN)>] [<ro: READONLY, RO, NOASSIGN>]  ;
                =>                                                      ;
                clsMkr:variable({ <(name1)> [, <(nameN)>] }, .f., !<.ro.>)

#command    INSTVAR <name1> [, <nameN>] [<ro: READONLY, RO, NOASSIGN>]  ;
                =>                                                      ;
                VAR <name1> [, <nameN>] <ro>

#xcommand   CLASS VAR <(name1)> [, <(nameN)>] [<ro: READONLY, RO, NOASSIGN>] ;
                =>                                                           ;
                clsMkr:variable({ <(name1)> [, <(nameN)>] }, .t., !<.ro.>)

#xcommand   CLASSVAR <*details*>   =>  CLASS VAR <details>

#command    MESSAGE <message1> [, <messageN>] [<ct: CONSTRUCTOR, CTOR>] ;
                =>                                                      ;
                clsMkr:method({ { <"message1">, { || <message1>() } }   ;
                             [, { <"messageN">, { || <messageN>() } }] }, <.ct.>, <.ct.>)

#command    MESSAGE <message> METHOD <method> [<ct: CONSTRUCTOR, CTOR>] ;
                =>                                                      ;
                clsMkr:method({{ <"message">, { || <method>() } }}, <.ct.>, <.ct.>)

#command    MESSAGE <message1> [, <messageN>] <def: DEFERRED, NULL>	;
                =>                                                      ;
                clsMkr:method({ { <"message1">, nil }                   ;
                             [, { <"messageN">, nil }] }, .f., .f.)

#xcommand   METHOD <method1> [, <methodN>] [<ct: CONSTRUCTOR, CTOR>]    ;
                =>                                                      ;
                MESSAGE <method1> [, <methodN>] <ct>

#xcommand   METHOD <msg1> = <mth1> [, <msgN> = <mthN>] [<ct: CONSTRUCTOR, CTOR>] ;
                =>                                                      ;
                MESSAGE <msg1> METHOD <mth1> <ct>                       ;
             [; MESSAGE <msgN> METHOD <mthN> <ct>]

#xcommand   METHOD <message> = <method> <ct: CONSTRUCTOR, CTOR>    ;
                =>                                                      ;
                MESSAGE <message> METHOD <method> <ct>

#xcommand   METHOD <method1> [, <methodN>] <def: DEFERRED, NULL>   ;
                =>                                                      ;
                MESSAGE <method1> [, <methodN>] <def>

#xcommand   CLASS MESSAGE <message1> [, <messageN>] ;
                =>                                  ;
                clsMkr:method({ { <"message1">, { || <message1>() } } ;
                             [, { <"messageN">, { || <messageN>() } }] }, .t., .f.)

#xcommand   CLASS MESSAGE <message> METHOD <method> ;
                =>                                  ;
                clsMkr:method({{ <"message">, { || <method>() } }}, .t., .f.)

#xcommand   CLASS MESSAGE <message1> [, <messageN>] <def: DEFERRED, NULL>   ;
                =>                                                          ;
                clsMkr:method({ { <"message1">, nil }                       ;
                             [, { <"messageN">, nil }] }, .t., .f.)

#command    CLASS METHOD <method1> [, <methodN>]   ;
                =>                                      ;
                CLASS MESSAGE <method1> [, <methodN>]

#command    CLASS METHOD <message> = <method>      ;
                =>                                      ;
                CLASS MESSAGE <message> METHOD <method>

#command    CLASS METHOD <method1> [, <methodN>] <def: DEFERRED, NULL> ;
                =>                                                          ;
                CLASS MESSAGE <method1> [, <methodN>] <def>

#command    _CYCTOR <name>([<params,...>])[()]              ;
                =>                                          ;
                __cyMethodType procedure <name> (<params>)

#command    _CYCTOR ([<params,...>])[()]                    ;
                =>                                          ;
                __cyMethodType procedure new (<params>)

#command    _CYSUPCTOR ([<params,...>])[()]                 ;
                =>                                          ;
                local __csyDummy := qself():super:new(<params>)

#command    _CYSUPCTOR <name>([<params,...>])[()]           ;
                =>                                          ;
                local __csyDummy := qself():super:<name>(<params>)

#command    CONSTRUCTOR <ctor> [, <superCtor> ]     ;
                =>                                  ;
                _CYCTOR <ctor>()                    ;;
                _CYSUPCTOR <superCtor>()            ;;
                local self := qself()

#command    METHOD PROCEDURE <name>             ;
                =>                              ;
                __cyMethodType procedure <name> ;;
                local self := qself()

#command    METHOD FUNCTION <name>              ;
                =>                              ;
                __cyMethodType function <name>  ;;
                local self := qself()

#define __cyMethodType

#translate  ::  =>  self:

#translate  @:<methodName>([<param1>] [, <paramN>]) ;
                =>                                  ;
                <methodName>(_cyPshSelf(<param1>) [, <paramN>])

// eof class(y).ch
