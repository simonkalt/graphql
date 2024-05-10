FILESYSL_TEXT   SEGMENT BYTE PUBLIC 'CODE'
        ASSUME cs:FILESYSL_TEXT, ds:NOTHING

        PUBLIC  __f_firstlfn
__f_firstlfn    PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si
                PUSH    ds

                MOV     cx, WORD PTR [bp + 0Eh]
                LDS     dx, DWORD PTR [bp + 06h]
                LES     di, DWORD PTR [bp + 0Ah]
                MOV     si, 1           ; DOS format date/time
                MOV     ax, 714Eh       ; Find first long filename
                INT     21h

                JNC     FirstNoError

                SUB     ax, ax

FirstNoError:   POP     ds
                POP     si
                POP     di
                POP     bp
                RETF
__f_firstlfn    ENDP

; здддддддддбддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддд©
; ЁFunction Ё__f_next ()                                                      Ё
; ЁPurpose  ЁFind next                                                        Ё
; юдддддддддаддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддды
        PUBLIC  __f_nextlfn
__f_nextlfn     PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si
                PUSH    ds

                MOV     bx, WORD PTR [bp + 06h]

                LES     di, DWORD PTR [bp + 08h]
                MOV     si, 1           ; DOS format date/time
                MOV     ax, 714Fh       ; Find next long filename
                INT     21h

                JNC     NextNoError

                SUB     ax, ax

NextNoError:    POP     ds
                POP     si
                POP     di
                POP     bp
                RETF
__f_nextlfn     ENDP

; здддддддддбддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддд©
; ЁFunction Ё__f_closelfn ()                                                  Ё
; ЁPurpose  ЁClose search                                                     Ё
; юдддддддддаддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддддды
        PUBLIC  __f_closelfn
__f_closelfn    PROC    FAR
                PUSH    bp
                MOV     bp, sp
                PUSH    di
                PUSH    si
                PUSH    ds

                MOV     bx, WORD PTR [bp + 06h]
                MOV     ax, 71A1h       ; Find close
                INT     21h

                POP     ds
                POP     si
                POP     di
                POP     bp
                RETF
__f_closelfn    ENDP

FILESYSL_TEXT   ENDS
                END
