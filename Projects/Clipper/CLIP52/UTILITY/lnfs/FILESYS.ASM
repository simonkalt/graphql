EXTRN           __upref:FAR             ; Check // parameters on Clipper cmd-line and environment
EXTRN           __errmsg:FAR            ; Display string to STDOUT

ifdef USE286
.286
endif

VERSION         EQU     '0.04á'

_DATA           SEGMENT WORD PUBLIC 'DATA'
        ; Public variables
        PUBLIC  __terror, __horror, __tversion, __txerror
        PUBLIC  __txclass, __txaction, __txlocus, __tlfn

__terror        DW      0               ; Last error code
__horror        DW      0               ; Last INT 24 error code
__tversion      DW      0               ; DOS version
__txerror       DW      0               ; Extended error
__txclass       DW      0               ; Extended error class
__txaction      DW      0               ; Extended error action
__txlocus       DW      0               ; Extended error locus
__tlfn          DW      0               ; Using long filenames?
nMask           DB      0
_DATA           ENDS

DGROUP  GROUP   _DATA

FILESYS_DATA    SEGMENT BYTE PUBLIC 'FAR_DATA'
cTemp           DB      256 DUP(0)
cClipTemp       DB      'CLIPTMPA.TMP', 0
cNOLFN          DB      'NOLFN', 0
cINFO           DB      'INFO', 0
cBanner         DB      'Dark Black Long File Name module v', VERSION, 13, 10, 0
cOffMessage     DB      '   (Using standard DOS filenames)', 13, 10, 0
cWin95Found     DB      '   (Win95 found - using long filenames)', 13, 10, 0
FILESYS_DATA    ENDS

FILESYS_TEXT    SEGMENT BYTE PUBLIC 'CODE'
        ASSUME cs:FILESYS_TEXT, ds:DGROUP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³ValidDrive ()                                                    ³
; ³Purpose  ³Check that a drive is valid                                      ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input:  AX - Drive number                                                  ³
; ³Output: AX - 0 = valid (no error)                                          ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
ValidDrive      PROC    NEAR
                CLC
                INT     11h             ; Get BIOS equipement list

                PUSH    ax
                AND     ax, 01h         ; Bit 0 == Floppy installed?
                POP     ax
                JZ      NoFloppy

                SUB     ah, ah

ifdef USE286
                SHR     al, 06h
else
                MOV     cl, 06h
                SHR     al, cl
endif
                CMP     dl, al
                JBE     RetZero

                CMP     dl, 01h
                JNZ     NoFloppy

                MOV     ax, -1
                JMP     SHORT ValidDriveRet

NoFloppy:       INC     dl
                MOV     ah, 36h
                INT     21h

                CMP     ax, -1
                JZ      ValidDriveRet

RetZero:        SUB     ax, ax

ValidDriveRet:  RET
ValidDrive      ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³LFNInit ()                                                       ³
; ³Purpose  ³Initialize long filename system                                  ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input:  Nothing                                                            ³
; ³Output: Nothing                                                            ³
; ³Effect: Sets __tlfn.  1 = LFN support, 0 = No LFN support                  ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
LFNInit         PROC    NEAR
ifdef USE286
                ENTER   2, 0
else
                PUSH    bp
                MOV     bp, sp
                SUB     sp, 2
endif
                PUSH    di
                PUSH    si

                MOV     WORD PTR [bp - 2], 0

                ; //INFO used?  If so, show our banner
ifdef USE286
                PUSH    SEG cINFO
                PUSH    OFFSET cINFO
else
                MOV     ax, SEG cINFO
                PUSH    ax
                MOV     ax, OFFSET cINFO
                PUSH    ax
endif
                CALL    __upref                         ; Check command line and env for //INFO
                ADD     sp, 4

                CMP     ax, -1
                JZ      NoBanner

ifdef USE286
                PUSH    SEG cBanner
                PUSH    OFFSET cBanner
else
                MOV     ax, SEG cBanner
                PUSH    ax
                MOV     ax, OFFSET cBanner
                PUSH    ax
endif
                CALL    __errmsg                        ; Display it
                ADD     sp, 4

                MOV     WORD PTR [bp - 2], 1

NoBanner:
                ; See if the //NOLFN parameter has been used
ifdef USE286
                PUSH    SEG cNOLFN
                PUSH    OFFSET cNOLFN
else
                MOV     ax, SEG cNOLFN
                PUSH    ax
                MOV     ax, OFFSET cNOLFN
                PUSH    ax
endif
                CALL    __upref
                ADD     sp, 4

                CMP     ax, -1
                JZ      LFNOK

                CMP     WORD PTR [bp - 2], 1
                JNE     DontShowDisabledMessage

ifdef USE286
                PUSH    SEG cOffMessage
                PUSH    OFFSET cOffMessage
else
                MOV     ax, SEG cOffMessage
                PUSH    ax
                MOV     ax, OFFSET cOffMessage
                PUSH    ax
endif
                CALL    __errmsg
                ADD     sp, 4

DontShowDisabledMessage:
                MOV     __tlfn, 0

                POP     si
                POP     di
ifdef USE286
                LEAVE
else
                MOV     sp, bp
                POP     bp
endif
                RET

LFNOK:
                PUSH    ds

                SUB     dx, dx
                MOV     si, OFFSET cTemp
                MOV     ax, SEG cTemp
                MOV     ds, ax
                MOV     ax, 7147h
                INT     21h

                POP     ds

                JC      LFNFail

                CMP     ax, 7100h
                JZ      LFNFail

                MOV     ax, 1
                JMP     SHORT LFNExit

LFNFail:        SUB     ax, ax

LFNExit:        MOV     __tlfn, ax

                CMP     WORD PTR [bp - 2], 1
                JNE     DontShowStatusMessage

                OR      ax, ax
                JZ      DontShowStatusMessage

ifdef USE286
                PUSH    SEG cWin95Found
                PUSH    OFFSET cWin95Found
else
                MOV     ax, SEG cWin95Found
                PUSH    ax
                MOV     ax, OFFSET cWin95Found
                PUSH    ax
endif
                CALL    __errmsg                ; Display lfn information message
                ADD     sp, 4


DontShowStatusMessage:

                POP     si
                POP     di
ifdef USE286
                LEAVE
else
                MOV     sp, bp
                POP     bp
endif
                RET
LFNInit         ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³ClearErrors ()                                                   ³
; ³Purpose  ³Clear error variables                                            ³
; ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
ClearErrors     PROC    NEAR
                SUB     ax, ax
                MOV     WORD PTR __terror, ax
                MOV     WORD PTR __horror, ax
                MOV     WORD PTR __txerror, ax
                RET
ClearErrors     ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tinit ()                                                       ³
; ³Purpose  ³Init                                                             ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Standard XNC entry point.                                                  ³
; ³Output: AX must be 0 on exit to inform startup that all is OK, otherwise   ³
; ³        Program startup aborts                                             ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC __tinit
__tinit         PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                MOV     __terror, 0

                MOV     ah, 30h         ; Get DOS version
                SUB     al, al
                INT     21h

                CMP     al, 0
                JNZ     initFail

                MOV     ax, 1           ; Default to 1.00

initFail:       MOV     bl, al
                MOV     bh, 00h
                MOV     ah, 64h
                MUL     ah
                ADD     ax, bx
                MOV     __tversion, ax

                CMP     ax, 012Ch
                JNC     $L5

                MOV     nMask, 07h
                JMP     SHORT $L7

$L5:            PUSH    ds
                MOV     ah, 25h                 ; Set interrupt handler
                MOV     al, 24h                 ; INT 24h
                MOV     dx, SEG FILESYS_TEXT    ; Segment of the handler
                MOV     ds, dx
                MOV     dx, OFFSET Int24Handler ; And it's offset
                INT     21h
                POP     ds

                MOV     nMask, 0FFh

$L7:            CALL    LFNInit

                SUB     ax, ax

                POP     si
                POP     di
                POP     bp
                RETF
__tinit         ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tend ()                                                        ³
; ³Purpose  ³Finish off                                                       ³
; ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __tend
__tend          PROC    FAR
                MOV     WORD PTR __terror, 0
                RETF
__tend          ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³GetExtendedErrorInfo ()                                          ³
; ³Purpose  ³Get extended error info for Open/Create/SetDevInfo               ³
; ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
GetExtendedErrorInfo            PROC    NEAR
                MOV     ax, WORD PTR __terror
                MOV     WORD PTR __txerror, ax

                SUB     ax, ax
                MOV     WORD PTR __txclass, ax
                MOV     WORD PTR __txaction, ax
                MOV     WORD PTR __txlocus, ax

                CMP     WORD PTR __tversion, 012Ch ; Can we get extended error information?
                JC      DontGetErrorInfo

                PUSH    ds
                MOV     ah, 59h                 ; Get extended error information
                MOV     bx, 0
                INT     21h
                POP     ds

                MOV     WORD PTR __txerror, ax
                SUB     ax, ax
                MOV     al, bh
                MOV     WORD PTR __txclass, ax
                MOV     al, bl
                MOV     WORD PTR __txaction, ax
                MOV     al, ch
                MOV     WORD PTR __txlocus, ax

DontGetErrorInfo: RET
GetExtendedErrorInfo            ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsCreate ()                                                    ³
; ³Purpose  ³Create file                                                      ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Inputs: DWORD PTR @ Stack + 06h: Filename                                  ³
; ³        WORD      @ Stack + 0Ah: Open mode                                 ³
; ³Output: AX = file handle or -1 for error (see __terror/_fsError for error) ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsCreate, _tcreat
_tcreat         PROC    FAR
_tcreat         ENDP
__fsCreate      PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                PUSH    ds

                CMP     __tlfn, 1
                JNE     ShortCreate

                ; Long Filename create
                MOV     ax, 716Ch       ; LFN Extended Open
                LDS     si, DWORD PTR [bp + 06h]
                MOV     di, 1
                MOV     bx, 0002h
                SUB     cx, cx
                MOV     dx, 0012h       ; Create, truncate if exists
                INT     21h

                JMP     SHORT ShortCreateRet

ShortCreate:
                ; Normal create
                MOV     ah, 3Ch
                LDS     dx, DWORD PTR [bp + 06h]
                MOV     cx, WORD PTR [bp + 0Ah]
                INT     21h

ShortCreateRet:
                POP     ds

                JNC     CreateOK

                MOV     WORD PTR __terror, ax
                CALL    GetExtendedErrorInfo

                MOV     ax, -1

CreateOK:       POP     si
                POP     di
                POP     bp
                RETF
__fsCreate      ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tctemp ()                                                      ³
; ³Purpose  ³Create temp                                                      ³
; ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __tctemp
__tctemp        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                CMP     WORD PTR __tversion, 0012Ch
                JNC     $L12

                PUSH    ds
                MOV     dx, SEG cTemp
                MOV     es, dx
                MOV     di, OFFSET cTemp
                LDS     si, DWORD PTR [bp + 06h]

TempLoop1:      LODSB
                STOSB
                OR      al, al
                JNZ     TempLoop1

                MOV     ds, dx          ; DX still holds our data SEG
                MOV     si, OFFSET cClipTemp
                INC     BYTE PTR [si + 07h]
                DEC     di

                POP     ds

TempLoop2:      LODSB
                STOSB
                OR      al, al
                JNZ     TempLoop2

                MOV     ah, 3Ch
                MOV     cx, WORD PTR [bp + 0Ah]
                MOV     dx, OFFSET cTemp
                INT     21h

                MOV     bx, ds
                JNC     $L17
                MOV     WORD PTR __terror, ax
                MOV     ax, -1
                JMP     SHORT TempRet

$L12:           PUSH    ds
                MOV     dx, ds
                MOV     es, dx
                MOV     di, OFFSET cTemp
                LDS     si, DWORD PTR [bp + 06h]

TempLoop3:      LODSB
                STOSB
                OR      al, al
                JNZ     TempLoop3

                MOV     ds, dx
                MOV     ah, 5Ah
                MOV     dx, OFFSET cTemp
                MOV     cx, WORD PTR [bp + 0Ah]
                INT     21h

                MOV     bx, ds
                POP     ds
                JNC     $L17

                MOV     WORD PTR __terror, ax
                MOV     ax, -1
                JMP     SHORT TempRet

$L17:           PUSH    ax
                PUSH    ds
                MOV     ds, bx
                MOV     si, dx
                LES     di, DWORD PTR [bp+00Ch]
                CLD

TempLoop4:      LODSB
                STOSB
                OR      al, al
                JNZ     TempLoop4

                POP     ds
                POP     ax

TempRet:        POP     si
                POP     di
                POP     bp
                RETF
__tctemp        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsOpen ()                                                      ³
; ³Purpose  ³Open                                                             ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Inputs: DWORD PTR @ Stack + 06h: Filename                                  ³
; ³        WORD @ Stack + 0Ah: Open mode                                      ³
; ³Output: AX = file handle or -1 for error (see __terror/_fsError for error) ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsOpen, _topen
_topen          PROC    FAR
_topen          ENDP
__fsOpen        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                PUSH    ds

                CMP     __tlfn, 1
                JNE     ShortOpen

                ; Long filename open
                MOV     ax, 716Ch
                MOV     bx, WORD PTR [bp + 0Ah]
                SUB     bh, bh
                SUB     cx, cx
                MOV     dx, 1
                LDS     si, DWORD PTR [bp + 06h]
                MOV     di, 1
                INT     21h

                JMP     SHORT ShortOpenRet
ShortOpen:
                ; Normal open
                MOV     ah, 3Dh
                MOV     al, BYTE PTR [bp + 0Ah]
                AND     al, BYTE PTR nMask
                LDS     dx, DWORD PTR [bp + 06h]
                INT     21h
ShortOpenRet:
                POP     ds

                JNC     $L21

                MOV     WORD PTR __terror, ax
                CALL    GetExtendedErrorInfo

                MOV     ax, -1

$L21:           POP     si
                POP     di
                POP     bp
                RETF
__fsOpen        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsClose ()                                                     ³
; ³Purpose  ³Close file                                                       ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: File handle                                     ³
; ³Output: AX = 0 for success, non 0 for error                                ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsClose, __tclose
__tclose        PROC    FAR
__tclose        ENDP
__fsClose       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 3Eh
                MOV     bx, WORD PTR [bp + 06h]
                INT     21h

                JNC     $L22

                MOV     WORD PTR __terror, ax

$L22:           POP     si
                POP     di
                POP     bp
                RETF
__fsClose       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsRead ()                                                      ³
; ³Purpose  ³Read from file                                                   ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD      @ Stack + 06h: File handle                                ³
; ³       DWORD PTR @ Stack + 08h: Read buffer                                ³
; ³       WORD      @ Stack + 0Ch: Number of bytes to read                    ³
; ³Output: AX = Number of bytes read from file.  If 0, then check for error   ³
; ³        or end of file                                                     ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsRead, __tread
__tread         PROC    FAR
__tread         ENDP
__fsRead        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                PUSH    ds
                MOV     ah, 3Fh
                MOV     bx, WORD PTR [bp + 06h]
                LDS     dx, DWORD PTR [bp + 08h]
                MOV     cx, WORD PTR [bp + 0Ch]
                INT     21h
                POP     ds

                JNC     $L23
                MOV     WORD PTR __terror, ax

                SUB     ax, ax

$L23:           POP     si
                POP     di
                POP     bp
                RETF
__fsRead        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsWrite ()                                                     ³
; ³Purpose  ³Write to file                                                    ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD      @ Stack + 06h: File handle                                ³
; ³       DWORD PTR @ Stack + 08h: Write buffer                               ³
; ³       WORD      @ Stack + 0Ch: Number of bytes to write                   ³
; ³Output: AX = Number of bytes written to file.  If 0, then check for error  ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsWrite, __twrite
__twrite        PROC    FAR
__twrite        ENDP
__fsWrite       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                PUSH    ds
                MOV     ah, 40h
                MOV     bx, WORD PTR [bp + 06h]
                LDS     dx, DWORD PTR [bp + 08h]
                MOV     cx, WORD PTR [bp + 0Ch]
                INT     21h

                POP     ds
                JNC     $L24

                MOV     WORD PTR __terror, ax

                SUB     ax, ax

$L24:           POP     si
                POP     di
                POP     bp
                RETF
__fsWrite       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsSeek ()                                                      ³
; ³Purpose  ³Seek                                                             ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD  @ Stack + 06h: File handle                                    ³
; ³       DWORD @ Stack + 08h: New file offset                                ³
; ³       WORD  @ Stack + 0Ch: Movement method                                ³
; ³Output: AX = New file offset                                               ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsSeek, __tlseek
__tlseek        PROC    FAR
__tlseek        ENDP
__fsSeek        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 42h
                MOV     bx, WORD PTR [bp + 06h]
                MOV     dx, WORD PTR [bp + 08h]
                MOV     cx, WORD PTR [bp + 0Ah]
                MOV     al, BYTE PTR [bp + 0Ch]
                INT     21h

                JNC     $L25
                MOV     WORD PTR __terror, ax

                SUB     ax, ax
                MOV     dx, ax

$L25:           POP     si
                POP     di
                POP     bp
                RETF
__fsSeek        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsLock ()                                                      ³
; ³Purpose  ³Lock                                                             ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD  @ Stack + 06h: File handle                                    ³
; ³       DWORD @ Stack + 08h: Start offset of lock                           ³
; ³       DWORD @ Stack + 0Ch: End offset of lock                             ³
; ³       WORD  @ Stack + 0Eh: Mode (0 = Lock, 1 = Unlock)                    ³
; ³Output: AX = New file offset                                               ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsLock, __tlock
__tlock         PROC    FAR
__tlock         ENDP
__fsLock        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 5Ch
                MOV     bx, WORD PTR [bp + 06h]
                MOV     dx, WORD PTR [bp + 08h]
                MOV     cx, WORD PTR [bp + 0Ah]
                MOV     di, WORD PTR [bp + 0Ch]
                MOV     si, WORD PTR [bp + 0Eh]
                MOV     al, BYTE PTR [bp + 10h]
                INT     21h
                JNC     $L26

                CMP     ax, 0021h
                MOV     bx, ax
                MOV     ax, 0
                JZ      $L27

$L26:           MOV     WORD PTR __terror, bx
                MOV     ax, 0001h

$L27:           POP     si
                POP     di
                POP     bp
                RETF
__fsLock        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tmname ()                                                      ³
; ³Purpose  ³Get machine name                                                 ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 06h: Buffer for machine name                    ³
; ³       (Warning: This buffer MUST be at least 16 bytes long)               ³
; ³Output: AX : 0 = Error (see _terror/_fsError)                              ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __tmname
__tmname        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                SUB     cx, cx
                CMP     WORD PTR __tversion, 0136h
                JC      $L28

                PUSH    ds
                MOV     ah, 5Eh
                MOV     al, 00h
                LDS     dx, DWORD PTR [bp + 06h]
                INT     21h
                POP     ds
                JNC     $L28

                MOV     WORD PTR __terror, ax
                MOV     cx, 00000h

$L28:           MOV     ax, cx

                POP     si
                POP     di
                POP     bp
                RETF
__tmname        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsDelete ()                                                    ³
; ³Purpose  ³Delete file                                                      ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 06h: Filename to delete                         ³
; ³Output: No return                                                          ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsDelete, _tdelete
_tdelete        PROC    FAR
_tdelete        ENDP
__fsDelete      PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                PUSH    ds

                CMP     __tlfn, 1
                JNE     ShortDelete

                ; Long filename delete
                MOV     ax, 7141h
                SUB     cx, cx
                MOV     si, cx
                LDS     dx, DWORD PTR [bp + 06h]
                INT     21h

                JMP     SHORT ShortDeleteRet

ShortDelete:
                ; Normal delete
                MOV     ah, 41h
                LDS     dx, DWORD PTR [bp + 06h]
                INT     21h

ShortDeleteRet:

                POP     ds
                JNC     $L29

                MOV     WORD PTR __terror, ax

$L29:           POP     si
                POP     di
                POP     bp
                RETF
__fsDelete      ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsRename ()                                                    ³
; ³Purpose  ³Rename file                                                      ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 06h: Filename to rename                         ³
; ³       DWORD PTR @ Stack + 0Ah: New filename                               ³
; ³Output: No return                                                          ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsRename, __trename
__trename       PROC    FAR
__trename       ENDP
__fsRename      PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 56h

                CMP     __tlfn, 1
                JNE     ShortRename

                MOV     ax, 7156h

ShortRename:    PUSH    ds
                LDS     dx, DWORD PTR [bp + 06h]
                LES     di, DWORD PTR [bp + 0Ah]
                INT     21h

                POP     ds
                JNC     $L30

                MOV     WORD PTR __terror, ax

$L30:           POP     si
                POP     di
                POP     bp
                RETF
__fsRename      ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__thcount ()                                                     ³
; ³Purpose  ³Set handle count                                                 ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: New file handle table size                      ³
; ³Output: AX = 0 for success, 1 for error (see _terror/_fsError)             ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __thcount
__thcount       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 30h
                INT     21h             ; Get DOS version number
                CMP     al, 0
                JNZ     $L31

                MOV     ax, 1

$L31:           MOV     bl, ah
                MOV     bh, 0
                MOV     ah, 64h
                MUL     ah
                ADD     ax, bx
                MOV     dx, 1
                CMP     ax, 014Ah
                JC      $L32

                MOV     ah, 67h         ; Set handle count
                MOV     bx, WORD PTR [bp + 06h]
                INT     21h
                JNC     $L32

                MOV     WORD PTR __terror, ax
                MOV     dx, 0

$L32:           MOV     ax, dx

                POP     si
                POP     di
                POP     bp
                RETF
__thcount       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsCommit ()                                                    ³
; ³Purpose  ³Commit/Flush file buffers                                        ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: File handle to flush                            ³
; ³Output: AX = Error code on error                                           ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsCommit, __tcommit
__tcommit       PROC    FAR
__tcommit       ENDP
__fsCommit      PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CMP     WORD PTR __tversion, 014Ah      ; Supported? (Checks DOS version for support)
                JC      $L33

                MOV     ah, 68h                         ; Flush file buffers
                MOV     bx, WORD PTR [bp + 06h]         ; File handle
                INT     21h

$L33:           POP     si
                POP     di
                POP     bp
                RETF
__fsCommit      ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__fsError ()                                                     ³
; ³Purpose  ³Get last error                                                   ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Output: AX = Last error code                                               ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsError, _terror
_terror         PROC    FAR
_terror         ENDP
__fsError       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                MOV     ax, WORD PTR __txerror
                CMP     ax, 0
                JZ      $L34

                CMP     ax, 0053h
                JNZ     $L35

                MOV     ax, WORD PTR __horror
                XOR     ah, ah
                ADD     ax, 0013h
                JMP     SHORT $L35

$L34:           MOV     ax, WORD PTR __terror

$L35:           POP     si
                POP     di
                POP     bp
                RETF
__fsError       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tisdevice ()                                                   ³
; ³Purpose  ³Does handle point to device?                                     ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: File handle to test                             ³
; ³Output: AX: 1 = Yes, 0 = No                                                ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __tisdevice
__tisdevice     PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                MOV     ax, 4400h
                MOV     bx, WORD PTR [bp + 06h]
                INT     21h
                JNC     $L36

                SUB     ax, ax

$L36:           AND     ax, 0080h
                JZ      $L37

                MOV     ax, 1

$L37:           POP     si
                POP     di
                POP     bp
                RETF
__tisdevice     ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tdevraw ()                                                     ³
; ³Purpose  ³Set device information                                           ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: File handle (must be character device)          ³
; ³       WORD @ Stack + 08h: Set or reset word? 0=reset, 1=set               ³
; ³Output: AX: 1 = Yes, 0 = No                                                ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
       PUBLIC  __tdevraw
__tdevraw       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     bx, WORD PTR [bp + 06h]
                XOR     dx, dx                  ; Remove all flags
                MOV     ax, WORD PTR [bp + 08h]
                CMP     ax, 0
                JZ      $L38

                OR      dl, 20h                 ; Set raw binary mode for handle 0 only

$L38:           CLC
                PUSH    ds
                MOV     ax, 4401h               ; Set device information
                INT     21h
                POP     ds

                MOV     ax, 1
                JNC     $L39

                MOV     WORD PTR __terror, ax
                CALL    GetExtendedErrorInfo
                SUB     ax, ax

$L39:           POP     si
                POP     di
                POP     bp
                RETF
__tdevraw       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__f_first ()                                                     ³
; ³Purpose  ³Find first                                                       ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 06h: Filespec string                            ³
; ³       DWORD PTR @ Stack + 0Ah: DTA buffer address (You don't need to know ³
; ³                        what a DTA is, just use the find_t structure... ;) ³
; ³       WORD      @ Stack + 0Eh: Attributes to include                      ³
; ³Output: AX: 1 = Found file, 0 = no files found (Quite annoyingly, this is  ³
; ³                the opposite of the standard C versions of this function)  ³
; ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³This is a standard version of this.  Because it effects the stack we can   ³
; ³only assume that 13 bytes exist for filename room.  Assuming this, we have ³
; ³special LFN versions of these functions.  Thus, any internal Clipper       ³
; ³functions using this will still be restricted to 8.3 filenames.            ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __f_first
__f_first       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si
                PUSH    ds

                LDS     dx, DWORD PTR [bp + 0Ah]
                MOV     ah, 1Ah         ; Set DTA address
                INT     21h

                MOV     cx, WORD PTR [bp + 0Eh]
                LDS     dx, DWORD PTR [bp + 06h]
                MOV     ah, 4Eh         ; Find first
                INT     21h

                MOV     ax, 1
                JNC     FindFirstRet

                XOR     ax, ax

FindFirstRet:   POP     ds
                POP     si
                POP     di
                POP     bp
                RETF
__f_first       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__f_next ()                                                      ³
; ³Purpose  ³Find next                                                        ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 0Ah: DTA buffer address                         ³
; ³Output: AX: 1 = Found file, 0 = no files found                             ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __f_next
__f_next        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si
                PUSH    ds

                LDS     dx, DWORD PTR [bp + 06h]
                MOV     ah, 4Fh        ; Find next
                INT     21h

                MOV     ax, 1
                JNC     FindNextRet

                XOR     ax, ax

FindNextRet:    POP     ds
                POP     si
                POP     di
                POP     bp
                RETF
__f_next        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__dspace ()                                                      ³
; ³Purpose  ³Fill diskspace structure                                         ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD      @ Stack + 0Ah: Drive number                               ³
; ³       DWORD PTR @ Stack + 0Ah: Pointer to dspace structure                ³
; ³Output: No return                                                          ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __dspace
__dspace        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                MOV     dx, WORD PTR [bp + 06h]
                LES     di, DWORD PTR [bp + 08h]
                MOV     ah, 36h
                INT     21h

                MOV     WORD PTR es:[di], bx
                MOV     WORD PTR es:[di + 02h], ax
                MOV     WORD PTR es:[di + 04h], dx
                MOV     WORD PTR es:[di + 06h], cx

                POP     si
                POP     di
                POP     bp
                RETF
__dspace        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tmkdir ()                                                      ³
; ³Purpose  ³MkDir                                                            ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 0Ah: Directory name to create                   ³
; ³Output: AX = 1 on success, 0 = failure (check _terror/_fsError)            ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsMkDir, __tmkdir
__fsMkDir       PROC    FAR
__fsMkDir       ENDP
__tmkdir        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 39h

                CMP     __tlfn, 1
                JNE     ShortMkDir

                MOV     ax, 7139h

ShortMkDir:     PUSH    ds
                LDS     dx, DWORD PTR [bp + 06h]
                INT     21h

                POP     ds
                JNC     $L42

                MOV     WORD PTR __terror, ax
                SUB     ax, ax
                JMP     SHORT $L43

$L42:           MOV     ax, 1

$L43:           POP     si
                POP     di
                POP     bp
                RETF
__tmkdir        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__trmdir ()                                                      ³
; ³Purpose  ³RmDir                                                            ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 0Ah: Directory name to remove                   ³
; ³Output: AX = 1 on success, 0 = failure (check _terror/_fsError)            ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsRmDir, __trmdir
__fsRmDir       PROC    FAR
__fsRmDir       ENDP
__trmdir        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 3Ah

                CMP     __tlfn, 1
                JNE     ShortRmDir

                MOV     ax, 713Ah

ShortRmDir:     PUSH    ds
                LDS     dx, DWORD PTR [bp + 06h]
                INT     21h

                POP     ds
                JNC     $L44

                MOV     WORD PTR __terror, ax
                MOV     ax, 0
                JMP     SHORT $L45

$L44:           MOV     ax, 1

$L45:           POP     si
                POP     di
                POP     bp
                RETF
__trmdir        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tchdir ()                                                      ³
; ³Purpose  ³Change directory                                                 ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: DWORD PTR @ Stack + 0Ah: Directory to change to                     ³
; ³Output: AX = 1 on success, 0 = failure (check _terror/_fsError)            ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsChDir, __tchdir
__fsChDir       PROC    FAR
__fsChDir       ENDP
__tchdir        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 3Bh

                CMP     __tlfn, 1
                JNE     ShortChDir

                MOV     ax, 713Bh

ShortChDir:     PUSH    ds
                LDS     dx, DWORD PTR [bp + 06h]
                INT     21h

                POP     ds
                JNC     $L46

                MOV     WORD PTR __terror, ax
                MOV     ax, 0
                JMP     SHORT $L47

$L46:           MOV     ax, 1

$L47:           POP     si
                POP     di
                POP     bp
                RETF
__tchdir        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tcurdrv ()                                                     ³
; ³Purpose  ³Get current drive                                                ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: Nothing                                                             ³
; ³Output: AX = Current drive number (0 = A:, 1 = B:, etc.)                   ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsCurDrv, __tcurdrv
__fsCurDrv      PROC    FAR
__fsCurDrv      ENDP
__tcurdrv       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 19h
                INT     21h

                MOV     ah, 0

                POP     si
                POP     di
                POP     bp
                RETF
__tcurdrv       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tchdrv ()                                                      ³
; ³Purpose  ³Change drive                                                     ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: New drive number (0 = A:, 1 = B:, etc.)         ³
; ³Output: AX = 0 for success, -1 for failure                                 ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsChDrv, __tchdrv
__fsChDrv       PROC    FAR
__fsChDrv       ENDP
__tchdrv        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     dl, BYTE PTR [bp + 06h]
                CALL    ValidDrive
                CMP     ax, 0FFFFh
                JZ      $L49

                MOV     dl, BYTE PTR [bp + 06h]
                MOV     ah, 0Eh
                INT     21h

                MOV     ax, 0

$L49:           POP     si
                POP     di
                POP     bp
                RETF
__tchdrv        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tisdrv ()                                                      ³
; ³Purpose  ³Is valid drive?                                                  ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: Drive number (0 = current drive!?, 1 = A:,      ³
; ³                          2 = B:, etc.) It uses the ValidDrive function in ³
; ³                          this module, this is why this parameter value is ³
; ³                          smegged up... <g>                                ³
; ³Output: AX: 0 = valid, -1 = invalid                                        ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsIsDrv, __tisdrv
__fsIsDrv       PROC    FAR
__fsIsDrv       ENDP
__tisdrv        PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     dl, BYTE PTR [bp + 06h]
                CALL    ValidDrive
                CMP     ax, 0FFFFh
                JZ      $L50

                MOV     ax, 0

$L50:           POP     si
                POP     di
                POP     bp
                RETF
__tisdrv        ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³__tcurdir ()                                                     ³
; ³Purpose  ³Get current directory                                            ³
; ÃÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
; ³Input: WORD @ Stack + 06h: Drive number (0 = current drive!?, 1 = A:,      ³
; ³                          2 = B:, etc.)                                    ³
; ³Output: DX:AX = pointer to current path                                    ³
; ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
        PUBLIC  __fsCurDir, __tcurdir
__fsCurDir      PROC    FAR
__fsCurDir      ENDP
__tcurdir       PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si

                CALL    SHORT ClearErrors

                MOV     ah, 47h

                CMP     __tlfn, 1
                JNE     ShortCurDir

                MOV     ax, 7147h

ShortCurDir:    PUSH    ds
                MOV     dx, SEG cTemp
                MOV     ds, dx
                MOV     si, OFFSET cTemp

                MOV     dx, WORD PTR [bp + 06h]         ; Drive

                MOV     ah, 47h
                INT     21h

                POP     ds

                MOV     dx, SEG cTemp                   ; Set return address
                MOV     ax, OFFSET cTemp

                POP     si
                POP     di
                POP     bp
                RETF
__tcurdir       ENDP

; ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
; ³Function ³Int24Handler ()                                                  ³
; ³Purpose  ³INT 24h interrupt handler                                        ³
; ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Int24Handler    PROC    FAR
                PUSH    ds
                PUSH    bx
                MOV     bx, DGROUP
                MOV     ds, bx
                MOV     bx, di
                MOV     bh, ah
                MOV     WORD PTR __horror, bx
                POP     bx
                POP     ds
                MOV     al, 03h         ; Fail the system call in progress
                IRET
Int24Handler    ENDP

FILESYS_TEXT    ENDS

                END
