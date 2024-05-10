_DATA           SEGMENT   WORD PUBLIC 'DATA'
                EXTRN   __tlfn:WORD
_DATA           ENDS

CONST           SEGMENT   WORD PUBLIC 'CONST'
CONST           ENDS

_BSS            SEGMENT   WORD PUBLIC 'BSS'
_BSS            ENDS

DGROUP          GROUP   CONST, _BSS, _DATA

                EXTRN   __retc:FAR
                EXTRN   __parc:FAR
                EXTRN   __parinfo:FAR
                EXTRN   _strcpy:FAR

                PUBLIC  LFNSHORT

CONVERT_TEXT    SEGMENT   WORD PUBLIC 'CODE'
                ASSUME  cs:CONVERT_TEXT
                ASSUME  ds:DGROUP

; здддддддддбддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддд©
; ЁFunction ЁLFNSHORT                                                         Ё
; ЁPurpose  ЁConvert long filename to it's short 8.3 equivelant               Ё
; юдддддддддаддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддды
LFNSHORT        PROC    FAR
ifdef USE286
                ENTER   110h, 0
else
                PUSH    bp
                MOV     bp, sp
                SUB     sp, 110h
endif
                PUSH    di
                PUSH    si

ifdef USE286
                PUSH    1
else
                MOV     ax, 1
                PUSH    ax
endif
                CALL    __parinfo
                ADD     sp, 2

                CMP     ax, 1                   ; Is it character?
                JE      ParameterOK

                MOV     WORD PTR [bp - 110h], 0
                JMP     SHORT ShortRet

ParameterOK:
                MOV     ax, 1
                PUSH    ax
                CALL    __parc
                ADD     sp, 2

                PUSH    ds
                MOV     ds, dx
                MOV     si, ax

                MOV     ax, ss
                MOV     es, ax
                LEA     di, WORD PTR [bp - 110h]

                CALL    FAR PTR LFNLongToShort

                POP     ds

ShortRet:
                PUSH    es
                PUSH    di
                CALL    __retc
                ADD     sp, 4

                POP     si
                POP     di
ifdef USE286
                LEAVE
else
                MOV     sp, bp
                POP     bp
endif
                RETF
LFNSHORT        ENDP

; здддддддддбддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддд©
; ЁFunction ЁLongToShort ()                                                   Ё
; ЁPurpose  ЁConvert long filename into short 8.3 version if same file        Ё
; цдддддддддаддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддд╢
; ЁInputs: DS:SI PTR - Filename to convert                                    Ё
; Ё        ES:DI PTR - Destination buffer for converted filename              Ё
; юддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддды
        PUBLIC  LFNLongToShort
LFNLongToShort  PROC    FAR
                PUSH    ds
                MOV     ax, DGROUP
                MOV     ds, ax
                CMP     __tlfn, 0
                POP     ds
                JZ      NoConversion

                MOV     ax, 7160h               ; Convert filename into 8.3 format (XXXXXX~N.XXX) to open it
                MOV     cx, 1
                INT     21h

                JNC     NoError

NoConversion:
                ; An error occured, so we copy the original filename into the target buffer

                PUSH    ds
                PUSH    si
                PUSH    es
                PUSH    di
                CALL    _strcpy
                POP     di
                POP     es
                POP     si
                POP     ds
NoError:

                RETF
LFNLongToShort  ENDP

CONVERT_TEXT    ENDS

                END
