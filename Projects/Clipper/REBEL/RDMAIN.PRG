#include "box.ch"
#include "sixcdx.ch"
#include "machsix.ch"
#include "set.ch"
#include "inkey.ch"

*******************************
**** Variable Assignments  ****
****                       ****
**** Bindley	      = BW ****
**** Kinray           = F  ****
**** H.D. Smith       = T  ****
**** Burlington/Amerisource/NC Mutual/Longs  = W  ****
**** Schein           = N  ****
**** Bergen-Brunswig  = B  ****
**** Rochester/Dakota = J  ****
**** Walsh            = BR ****
**** D & K Drug       = BE ****
**** Walgreens        = WG ****
**** McKesson         = MC ****
**** Cardinal         = CL ****
**** F. Dohmen/Smith  = FD ****
**** Harvard Drug     = H  ****
**** Bellco/M&D       = BD ****
****                       ****
*******************************

LOCAL MITEM[8]

SET EXCLU OFF

PUBLIC strHarvardVendorNumber := "01747", _SSCREEN
PUBLIC DISKDIR := DISKNAME()+":"+DIRNAME()
PRIVATE SUCCESS := .F.

MITEM[1]='Maintain Inventory Master File'
MITEM[2]='Update Customer Transmittal Data'
MITEM[3]='Enter Bergen Orders'
MITEM[4]='Enter McKesson Orders'
MITEM[5]='Print Inventory Reports'
MITEM[6]='Prepare Data For Transmission'
MITEM[7]='Perform System Maintenance'
MITEM[8]='Exit Inventory Transmittal System'

INITIALSCREEN()

RDDSETDEFAULT("SIXCDX")

IF !NETUSE("RDINV",.F.,1,"RDINV")
  MSGBOX("SYSTEM IS BUSY...")
ENDIF

NETUSE("RDINV",.F.,0,"RDINV")
SET ORDER TO TAG RDINV
GO TOP

DO NTX
INKEY(1.5)

SET DELETED ON
MCHOICE=1
DO WHILE .T.
  DO BORDER WITH .T.
  @ 5,6 SAY "Main Menu...                   "
  @ 7,6 CLEAR TO 22,44
  @ 6,5 TO 23,45 DOUBLE
  @ 15,47 SAY "ENTER = SELECT"
  @ 16,47 SAY "ESC   = QUIT  "
  MCHOICE=ACHOICE(10,6,22,44,MITEM,.T.,'',MCHOICE)
  IF LASTKEY()=27.OR.MCHOICE=8
    CLEAR
    RETURN
  ENDIF
  @ 14,47 CLEAR TO 17,60
  DO CASE
    CASE MCHOICE = 1
      DO RDINV
    CASE MCHOICE = 2
      DO RDCUST
    CASE MCHOICE = 3
      DO RDORDER
    CASE MCHOICE = 4
      DO MCORDER
    CASE MCHOICE = 5
      DO RDREPT
    CASE MCHOICE = 6
      DO RDTRANS
    CASE MCHOICE = 7
      @ 16,47 SAY "<--SELECTED"
      MYORN=.F.
      @ 23,0 SAY 'Continue with System Maintenance? [y/n]' GET MYORN PICT 'Y'
      READ
      IF MYORN
        @ 23,0 CLEAR
        @ 23,0 SAY "Rebuilding.  Please Wait..."
        SELECT RDINV
        USE
        RUN DEL *.CDX > NUL
        DO NTX
        @ 23,0 CLEAR
        @ 23,0 SAY "System Maintenance Completed.  Press Any Key..."
        INKEY(0)
      ENDIF
  ENDCASE
ENDDO
RETURN

PROCEDURE RDORDER
  DO WHILE .T.
    DO BORDER WITH .T.
    @ 7,5 TO 19,40 DOUBLE
    SET CONFIRM ON
    @ 6,5 SAY "Enter Bergen-Brunswig Orders..."
    MPO=SPACE(8)
    @ 10,6 SAY "Purchase Order:" GET MPO PICTURE "!!!!!!!!"
    READ
    @ 9,6 CLEAR TO 18,44
    IF LASTKEY()=27
    RETURN
  ENDIF
  SELECT RDINV
  SET ORDER TO TAG RDBITEM
  GO TOP
  MFNAME=DISKDIR+"\PO\"+TRIM(MPO)+'.BPO'
  MINAME=DISKDIR+"\PO\"+TRIM(MPO)+'.IDX'
  MYORN=.F.
  IF FILE(MFNAME)
    MYORN =.T.
      USE &MFNAME ALIAS TEMP NEW
      IF FILE(MINAME)
        SET INDEX TO (MINAME)
      ENDIF
    ELSE
      @ 23,0 CLEAR
      @ 23,0 SAY 'Purchase order not on file.  Create new order? (y/n)' GET MYORN PICT "y"
      READ
      @ 23,0 CLEAR
      IF MYORN
        COPY STRUCTURE TO &MFNAME
        USE &MFNAME ALIAS TEMP NEW
        INDEX ON BITEMNO TO (MINAME)
      ENDIF
    ENDIF
    IF .NOT.MYORN
      IF ALIAS() # "RDINV"
        USE
      ENDIF
      SELECT RDINV
      GO TOP
      LOOP
    ENDIF
    @ 23,0 CLEAR
    @ 23,0 SAY 'Loading purchase order.  Please wait...'
    IF !FILE(MINAME)
      INDEX ON BITEMNO TO (MINAME)
    ENDIF
    SET INDEX TO (MINAME)
    @ 23,0 CLEAR
    DECLARE FLD[6],PIC[6],HEAD[6]
    FLD[1] = 'BITEMNO'
    FLD[2] = 'RITEM'
    FLD[3] = 'RNDC'
    FLD[4] = 'BQTYORD'
    FLD[5] = 'BPRICE'
    HEAD[1] = 'Bergen;Item #'
    HEAD[2] = 'Rebel;Item #'
    HEAD[3] = 'NDC #'
    HEAD[4] = 'Qty;Order'
    HEAD[5] = 'Price'
    STORE space(30) TO MRITEM
    STORE SPACE(10) TO MRNDC
    STORE SPACE(6) TO MBITEMNO
    STORE 0 TO MBQTYORD,MBPRICE
    DO BORDER WITH .T.
    @ 4,2 TO 14,77 DOUBLE
    @ 15,2 SAY "FIND    = ANY CHARACTER"
    @ 16,2 SAY "ADD NEW = INSERT       "
    @ 17,2 SAY "CHANGE  = ENTER        "
    @ 18,2 SAY "DELETE  = DELETE       "
    @ 19,2 SAY "EXIT    = ESC          "
    KEYBOARD CHR(19)
    DONE = .F.
    DO WHILE .NOT.DONE
      DBEDIT(5,3,13,76,FLD,"BORDFUNC",PIC,HEAD)
    ENDDO
    USE
    SELECT RDINV
    GO TOP
  ENDDO
RETURN

PROCEDURE MCORDER
  DO WHILE .T.
    DO BORDER WITH .T.
    @ 7,5 TO 19,45 DOUBLE
    SET CONFIRM ON
    @ 6,5 SAY "Enter McKesson Orders..."
    MPO=SPACE(8)
    @ 10,6 SAY "Purchase Order:" GET MPO PICTURE "!!!!!!!"
    READ
    @ 9,6 CLEAR TO 18,44
    IF LASTKEY()=27
      RETURN
    ENDIF
    SELECT RDINV
    SET ORDER TO TAG RDMCITEM
    GO TOP
    MFNAME=DISKDIR+"\PO\"+TRIM(MPO)+'.MPO'
    MINAME=DISKDIR+"\PO\M"+TRIM(MPO)+'.IDX'
    MYORN=.F.
    IF FILE(MFNAME)
      MYORN =.T.
      USE &MFNAME ALIAS TEMP NEW
      IF FILE(MINAME)
        SET INDEX TO (MINAME)
      ENDIF
    ELSE
      @ 23,0 CLEAR
      @ 23,0 SAY 'Purchase order not on file.  Create new order? (y/n)' GET MYORN PICT "y"
      READ
      @ 23,0 CLEAR
      IF MYORN
        COPY STRUCTURE TO &MFNAME
        USE &MFNAME ALIAS TEMP NEW
        INDEX ON MCITEMNO TO (MINAME)
      ENDIF
    ENDIF
    IF .NOT.MYORN
      IF ALIAS() # "RDINV"
        USE
      ENDIF
      SELECT RDINV
      GO TOP
      LOOP
    ENDIF
    @ 23,0 CLEAR
    @ 23,0 SAY 'Loading purchase order.  Please wait...'
    IF !FILE(MINAME)
      INDEX ON MCITEMNO TO (MINAME)
    ENDIF
    SET INDEX TO (MINAME)
    @ 23,0 CLEAR
    DECLARE FLD[6],PIC[6],HEAD[6]
    FLD[1] = 'MCITEMNO'
    FLD[2] = 'RITEM'
    FLD[3] = 'RNDC'
    FLD[4] = 'MCQTYAVAIL'
    FLD[5] = 'MCPRICE'
    HEAD[1] = 'McKesson;Item #'
    HEAD[2] = 'Rebel;Item #'
    HEAD[3] = 'NDC #'
    HEAD[4] = 'Qty;Order'
    HEAD[5] = 'Price'
    STORE space(30) TO MRITEM
    STORE SPACE(10) TO MRNDC
    STORE SPACE(7) TO MMCITEMNO
    STORE 0 TO MMCQTYAVAIL,MMCPRICE
    DO BORDER WITH .T.
    @ 4,2 TO 14,77 DOUBLE
    @ 15,2 SAY "FIND    = ANY CHARACTER"
    @ 16,2 SAY "ADD NEW = INSERT       "
    @ 17,2 SAY "CHANGE  = ENTER        "
    @ 18,2 SAY "DELETE  = DELETE       "
    @ 19,2 SAY "EXIT    = ESC          "
    KEYBOARD CHR(19)
    DONE = .F.
    DO WHILE .NOT.DONE
      DBEDIT(5,3,13,76,FLD,"MCORDFUNC",PIC,HEAD)
    ENDDO
    USE
    SELECT RDINV
    GO TOP
  ENDDO
RETURN

FUNCTION BORDFUNC
  PARAMETER STATUS, FLD_PTR

  RETVAL = 1
  MLAST = LASTKEY()
  DO CASE
    CASE MLAST = 27                              && ESCAPE KEY IS PRESSED
      DONE  = .T.
      RETVAL = 0
    CASE MLAST = 22                              && INSERT KEY IS PRESSED
      SAVE SCREEN TO TEMP
        DO ORDINIT
        DO ORDGET WITH "A"
        READ
        IF LASTKEY()<>27
          ADDREC( 0 )
          IF RECLOCK(2)
            DO ORDREPL
          ELSE
            MSGBOX("Record is in use.  Try again later.")
          ENDIF
        ENDIF
        KEYBOARD CHR(19)+CHR(19)+CHR(19)
        RESTORE SCREEN FROM TEMP
        KEYBOARD CHR(19)
        UNLOCK
      RETVAL = 1
    CASE MLAST =  7                              && DELETE KEY IS PRESSED
      SAVE SCREEN TO TEMP
      DO ORDSTORE
      DO ORDGET WITH 'D'
      CLEAR GETS
      @ 23,0 CLEAR
      MYORN = .F.
      @ 23,0 SAY "Are you sure you want to DELETE?" GET MYORN PICT 'Y'
      READ
      IF MYORN
        IF RECLOCK(3)
          REPLACE RITEM WITH "*** DELETED ***"
          DELETE
          UNLOCK
        ELSE
          MSGBOX("Record is in use.  Try again later.")
        ENDIF
      ENDIF
      RESTORE SCREEN FROM TEMP
      RETVAL = 1
    CASE MLAST = 13                              && ENTER KEY IS PRESSED
      IF BITEMNO<>SPACE(6)
        IF RECLOCK(3)
          SAVE SCREEN TO TEMP
          DO ORDSTORE
          DO ORDGET WITH 'E'
          READ
          IF LASTKEY()<>27
            DO ORDREPL
          ENDIF
          RESTORE SCREEN FROM TEMP
          UNLOCK
        ELSE
          MSGBOX("Record is in use.  Try again later.")
        ENDIF
      ENDIF
      RETVAL = 1
    CASE (MLAST>=48 .AND.MLAST<=57)
      SET CURSOR ON
      R = RECNO()
      MCHAR = CHR(MLAST)+SPACE(29)
      KEYBOARD CHR(4)
      @ 23,0 CLEAR
      @ 23,0 SAY "Find: " GET MCHAR PICTURE "@!"
      READ
      @ 23,0 CLEAR
      SEEK TRIM(MCHAR)
      IF FOUND()
        R=RECNO()
      ENDIF
      GOTO R
      KEYBOARD CHR(19)
      SET CURSOR OFF
      RETVAL = 0
  ENDCASE
RETURN RETVAL


FUNCTION MCORDFUNC
  PARAMETER STATUS, FLD_PTR

  RETVAL = 1
  MLAST = LASTKEY()
  DO CASE
    CASE MLAST = 27                              && ESCAPE KEY IS PRESSED
      DONE  = .T.
      RETVAL = 0
    CASE MLAST = 22                              && INSERT KEY IS PRESSED
      SAVE SCREEN TO TEMP
        DO MCORDINIT
        DO MCORDGET WITH "A"
        READ
        IF LASTKEY()<>27
          ADDREC( 0 )
          IF RECLOCK(2)
            DO MCORDREPL
          ELSE
            MSGBOX("Record is in use.  Try again later.")
          ENDIF
        ENDIF
        KEYBOARD CHR(19)+CHR(19)+CHR(19)
        RESTORE SCREEN FROM TEMP
        KEYBOARD CHR(19)
        UNLOCK
      RETVAL = 1
    CASE MLAST =  7                              && DELETE KEY IS PRESSED
      SAVE SCREEN TO TEMP
      DO MCORDSTORE
      DO MCORDGET WITH 'D'
      CLEAR GETS
      @ 23,0 CLEAR
      MYORN = .F.
      @ 23,0 SAY "Are you sure you want to DELETE?" GET MYORN PICT 'Y'
      READ
      IF MYORN
        IF RECLOCK(3)
          REPLACE RITEM WITH "*** DELETED ***"
          DELETE
          UNLOCK
        ELSE
          MSGBOX("Record is in use.  Try again later.")
        ENDIF
      ENDIF
      RESTORE SCREEN FROM TEMP
      RETVAL = 1
    CASE MLAST = 13                              && ENTER KEY IS PRESSED
      IF MCITEMNO<>SPACE(6)
        IF RECLOCK(3)
          SAVE SCREEN TO TEMP
          DO MCORDSTORE
          DO MCORDGET WITH 'E'
          READ
          IF LASTKEY()<>27
            DO MCORDREPL
          ENDIF
          RESTORE SCREEN FROM TEMP
          UNLOCK
        ELSE
          MSGBOX("Record is in use.  Try again later.")
        ENDIF
      ENDIF
      RETVAL = 1
    CASE (MLAST>=48 .AND.MLAST<=57)
      SET CURSOR ON
      R = RECNO()
      MCHAR = CHR(MLAST)+SPACE(29)
      KEYBOARD CHR(4)
      @ 23,0 CLEAR
      @ 23,0 SAY "Find: " GET MCHAR PICTURE "@!"
      READ
      @ 23,0 CLEAR
      SEEK TRIM(MCHAR)
      IF FOUND()
        R=RECNO()
      ENDIF
      GOTO R
      KEYBOARD CHR(19)
      SET CURSOR OFF
      RETVAL = 0
  ENDCASE
RETURN RETVAL


PROCEDURE ORDGET
PARAMETER MODE
  SET CURSOR ON
  SET CONFIRM ON
  @ 14, 40  TO 20, 77    DOUBLE
  @ 15, 41  CLEAR TO 19,76
  DO CASE
    CASE MODE = 'A'
      @ 15, 41  SAY "Add Item:"
    CASE MODE = 'D'
      @ 15, 41  SAY "Delete Item:"
      KEYBOARD CHR(27)
    CASE MODE = 'E'
      @ 15, 41  SAY "Change Item:"
  ENDCASE
  @ 16,43 SAY "Item    :" GET  MBITEMNO VALID OKBITEM(MBITEMNO)
  @ 17,43 SAY "Quantity:" GET  MBQTYORD PICT "9999"
  *@ 17,60 SAY "Price:" GET MBPRICE PICT "9999.99" WHEN .F.
  @ 18,43 SAY RITEM
  @ 19,43 SAY RNDC
  READ
RETURN

PROCEDURE MCORDGET
PARAMETER MODE
  SET CURSOR ON
  SET CONFIRM ON
  @ 14, 40  TO 20, 77    DOUBLE
  @ 15, 41  CLEAR TO 19,76
  DO CASE
    CASE MODE = 'A'
      @ 15, 41  SAY "Add Item:"
    CASE MODE = 'D'
      @ 15, 41  SAY "Delete Item:"
      KEYBOARD CHR(27)
    CASE MODE = 'E'
      @ 15, 41  SAY "Change Item:"
  ENDCASE
  @ 16,43 SAY "Item    :" GET  MMCITEMNO VALID OKMCITEM(MMCITEMNO)
  @ 17,43 SAY "Quantity:" GET  MMCQTYAVAIL PICT "9999"
  @ 18,43 SAY RITEM
  @ 19,43 SAY RNDC
  READ
RETURN


PROCEDURE OKBITEM
PARAMETER MSEEKVAR
  MRETVAL= .F.
  SELECT RDINV
  SEEK MSEEKVAR
  IF .NOT.EOF()
    MRETVAL=.T.
    @ 18,43 SAY RITEM
    @ 19,43 SAY RNDC
    STORE RITEM TO MRITEM
    STORE RNDC TO MRNDC
    STORE BPRICE TO MBPRICE
  ELSE
    ??CHR(7)
    @ 23,0 SAY "Item not in master file.  Press any key to re-enter..."
    INKEY(0)
    @ 23,0 CLEAR
  ENDIF
  SELECT TEMP
RETURN MRETVAL

PROCEDURE OKMCITEM
PARAMETER MSEEKVAR
  MRETVAL= .F.
  SELECT RDINV
  SEEK MSEEKVAR
  IF .NOT.EOF()
    MRETVAL=.T.
    @ 18,43 SAY RITEM
    @ 19,43 SAY RNDC
    STORE RITEM TO MRITEM
    STORE RNDC TO MRNDC
    STORE MCPRICE TO MMCPRICE
  ELSE
    ??CHR(7)
    @ 23,0 SAY "Item not in master file.  Press any key to re-enter..."
    INKEY(0)
    @ 23,0 CLEAR
  ENDIF
  SELECT TEMP
RETURN MRETVAL


PROCEDURE ORDINIT
  STORE space(30) TO MRITEM
  STORE space(10) TO MRNDC
  STORE SPACE(6) TO MBITEMNO
  STORE 0 TO MBQTYORD,MBPRICE
RETURN

PROCEDURE MCORDINIT
  STORE space(30) TO MRITEM
  STORE space(10) TO MRNDC
  STORE SPACE(7) TO MMCITEMNO
  STORE 0 TO MMCQTYAVAIL,MMCPRICE
RETURN

PROCEDURE ORDSTORE
  STORE RITEM TO MRITEM
  STORE RNDC TO MRNDC
  STORE BITEMNO TO MBITEMNO
  STORE BQTYORD TO MBQTYORD
  STORE BPRICE TO MBPRICE
RETURN

PROCEDURE MCORDSTORE
  STORE RITEM TO MRITEM
  STORE RNDC TO MRNDC
  STORE MCITEMNO TO MMCITEMNO
  STORE MCQTYAVAIL TO MMCQTYAVAIL
  STORE MCPRICE TO MMCPRICE
RETURN

PROCEDURE ORDREPL
  REPLACE RITEM WITH MRITEM
  REPLACE RNDC WITH MRNDC
  REPLACE BITEMNO WITH MBITEMNO
  REPLACE BQTYORD WITH MBQTYORD
  REPLACE BPRICE WITH MBPRICE
RETURN

PROCEDURE MCORDREPL
  REPLACE RITEM WITH MRITEM
  REPLACE RNDC WITH MRNDC
  REPLACE MCITEMNO WITH MMCITEMNO
  REPLACE MCQTYAVAIL WITH MMCQTYAVAIL
  REPLACE MCPRICE  WITH MMCPRICE
RETURN


PROCEDURE RDINV
  SELECT RDINV
  GO TOP
  @ 5,5 SAY "Maintain Inventory Master File..."
  DECLARE FLD[6],PIC[6],HEAD[6]
  FLD[1] = 'RITEM'
  FLD[2] = 'RNDC'
  HEAD[1] = 'Item'
  HEAD[2] = 'NDC #'
  STORE space(30) TO MRITEM
  STORE SPACE(10) TO MRNDC
  @ 12,47 SAY "FIND    = ANY CHARACTER"
  @ 13,47 SAY "ADD NEW = INSERT       "
  @ 14,47 SAY "CHANGE  = ENTER        "
  @ 15,47 SAY "DELETE  = DELETE       "
  @ 16,47 SAY "EXIT    = ESC          "
  KEYBOARD CHR(19)
  DONE = .F.
  DO WHILE .NOT. DONE
    DBEDIT(9,6,19,39,FLD,"RITEMFUNC",PIC,HEAD)
  ENDDO
  RELEASE ALL
  *CLOS ALL
RETURN


FUNCTION RITEMFUNC
  PARAMETER STATUS, FLD_PTR

  RETVAL = 1
  MLAST = LASTKEY()
  DO CASE
    CASE MLAST = 27                              && ESCAPE KEY IS PRESSED
      DONE  = .T.
      RETVAL = 0
    CASE MLAST = 22                              && INSERT KEY IS PRESSED
      SAVE SCREEN TO TEMP
      DO INVINIT
      DO INVGET WITH "A"
      READ
      IF LASTKEY()<>27
        ADDREC( 0 )
        DO INVREPL
      ENDIF
      KEYBOARD CHR(19)+CHR(19)+CHR(19)
      RESTORE SCREEN FROM TEMP
      KEYBOARD CHR(19)
      RETVAL = 1
    CASE MLAST =  7                              && DELETE KEY IS PRESSED
      SAVE SCREEN TO TEMP
      DO INVSTORE
      DO INVGET WITH 'D'
      CLEAR GETS
      @ 23,0 CLEAR
      MYORN = .F.
      @ 23,0 SAY "Are you sure you want to DELETE?" GET MYORN PICT 'Y'
      READ
      IF MYORN
        IF RECLOCK(3)
          REPLACE RITEM WITH "*** DELETED ***"
          DELETE
          UNLOCK
        ELSE
          MSGBOX("Record is in use.  Try again later.")
        ENDIF
      ENDIF
      RESTORE SCREEN FROM TEMP
      RETVAL = 1
    CASE MLAST = 13                              && ENTER KEY IS PRESSED
      IF RECLOCK(3)
        SAVE SCREEN TO TEMP
        DO INVSTORE
        DO INVGET WITH 'E'
        READ
        IF LASTKEY()<>27
          DO INVREPL
        ENDIF
        RESTORE SCREEN FROM TEMP
        UNLOCK
      ELSE
        MSGBOX("Record is in use.  Try again later.")
      ENDIF
      RETVAL = 1
    CASE (MLAST>=49.AND.MLAST<=122)
      SET CURSOR ON
      R = RECNO()
      MCHAR = CHR(MLAST)+SPACE(29)
      KEYBOARD CHR(4)
      @ 23,0 CLEAR
      @ 23,0 SAY "Find: " GET MCHAR PICTURE "@!"
      READ
      @ 23,0 CLEAR
      SEEK TRIM(MCHAR)
      IF FOUND()
        R=RECNO()
      ENDIF
      GOTO R
      KEYBOARD CHR(19)
      SET CURSOR OFF
      RETVAL = 0
  ENDCASE
RETURN RETVAL

PROCEDURE INVGET
PARAMETER MODE
  SET CURSOR ON
  SET CONFIRM ON
  @ 10, 20  TO 17, 62    DOUBLE
  @ 11, 21  CLEAR TO 16,61
  DO CASE
    CASE MODE = 'A'
      @ 11, 22  SAY "Enter New Item:"
    CASE MODE = 'D'
      @ 11, 22  SAY "Confirm Deletion:"
    CASE MODE = 'E'
      @ 11, 22  SAY "Type Changes:"
  ENDCASE
  @ 14, 22 SAY "Item :" GET  MRITEM
  @ 15, 22 SAY "NDC #:" GET  MRNDC
RETURN

PROCEDURE INVINIT
  STORE space(30) TO MRITEM
  STORE space(10) TO MRNDC
RETURN

PROCEDURE INVSTORE
  STORE RITEM TO MRITEM
  STORE RNDC TO MRNDC
RETURN

PROCEDURE INVREPL
    REPLACE RITEM WITH MRITEM
    REPLACE RNDC WITH MRNDC
RETURN


PROCEDURE RDCUST
  LOCAL MCUST[16]
  * MCUST[1]='Bindley Western'
  * MCUST[2]='Kinray'
  * MCUST[3]='H.D. Smith'
  * MCUST[4]='Burlington/Amerisource'
  * MCUST[5]='Schein'
  * MCUST[6]='Bergen-Brunswig'
  * MCUST[7]='Rochester/Dakota'
  * MCUST[8]='Walsh'
  * MCUST[9]='D & K Drug'
  * MCUST[10]='Walgreens'
  * MCUST[11]='McKesson'
  * MCUST[12]='Cardinal Health'
  * MCUST[13]='F. Dohmen/Smith Drug'
  * MCUST[14]='Harvard Drug'
  * MCUST[15]='Bellco Drug/M&D'
  
  MCUST[1]='Bindley Western'
  MCUST[2]='Walsh'
  MCUST[3]='Bellco Drug/M&D'
  MCUST[4]='Bergen-Brunswig'
  MCUST[5]='Burlington/Amerisource/NC Mutual/Longs'
  MCUST[6]='Cardinal Health'
  MCUST[7]='D & K Drug'
  MCUST[8]='F. Dohmen/Smith Drug'
  MCUST[9]='H.D. Smith'
  MCUST[10]='Harvard Drug'
  MCUST[11]='Kinray'
  MCUST[12]='McKesson'
  MCUST[13]='Schein'
  MCUST[14]='Rochester/Dakota'
  MCUST[15]='Walgreens'
  MCUST[16]='Return to Main Menu'
  MCUSTNO=1

  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP
  @ 7,6 CLEAR TO 22,44
  @ 5,6 SAY 'Select Customer to Update...'
  @ 15,47 SAY "ENTER = SELECT"
  @ 16,47 SAY "ESC   = QUIT  "
  MCUSTNO=ACHOICE(7,6,22,44,MCUST,.T.,'',MCUSTNO)
  IF LASTKEY()=27.OR.MCUSTNO=16
    RETURN
  ENDIF
  
  DO CASE
	CASE MCUSTNO = 1
		MCUSTNO = 1
	CASE MCUSTNO = 2
		MCUSTNO = 8
	CASE MCUSTNO = 3
		MCUSTNO = 15
	CASE MCUSTNO = 4
		MCUSTNO = 6
	CASE MCUSTNO = 5
		MCUSTNO = 4
	CASE MCUSTNO = 6
		MCUSTNO = 12
	CASE MCUSTNO = 7
		MCUSTNO = 9
	CASE MCUSTNO = 8
		MCUSTNO = 13
	CASE MCUSTNO = 9
		MCUSTNO = 3
	CASE MCUSTNO = 10
		MCUSTNO = 14
	CASE MCUSTNO = 11
		MCUSTNO = 2
	CASE MCUSTNO = 12
		MCUSTNO = 11
	CASE MCUSTNO = 13
		MCUSTNO = 5
	CASE MCUSTNO = 14
		MCUSTNO = 7
	CASE MCUSTNO = 15
		MCUSTNO = 10
  ENDCASE
  
  * @ 5,5 SAY 'Update Customer Transmittal Data...'
  DECLARE FLD[8],PIC[8],HEAD[8]

  STORE space(30) TO MRITEM
  STORE space(6)  TO MFITEMNO
  STORE 0         TO MFPRICE
  STORE 0         TO MFQTYAVAIL
  STORE 0         TO MFCASEPACK
  STORE .F.       TO MFTRANSMIT
  STORE .F.       TO MFVALID
  STORE 0         TO MBITEMNO
  STORE 0         TO MBQTYAVAIL
  STORE 0         TO MBPRICE
  STORE 0         TO MBCASEPACK
  STORE 0         TO MBMINORDER
  STORE .F.       TO MBTRANSMIT
  STORE SPACE(13) TO MTITEMNO
  STORE 0         TO MTQTYAVAIL
  STORE 0         TO MTPRICE
  STORE 0         TO MTCASEPACK
  STORE .F.       TO MTTRANSMIT
  STORE SPACE(13) TO MWITEMNO
  STORE 0         TO MWQTYAVAIL
  STORE 0         TO MWPRICE
  STORE 0         TO MWCASEPACK
  STORE .F.       TO MWTRANSMIT
  STORE SPACE(13) TO MNITEMNO
  STORE 0         TO MNQTYAVAIL
  STORE 0         TO MNPRICE
  STORE 0         TO MNCASEPACK
  STORE .F.       TO MNTRANSMIT
  STORE SPACE(6)  TO MBWIN
  STORE SPACE(13) TO MBWITEMNO
  STORE 0         TO MBWQTYAVAIL
  STORE 0         TO MBWPRICE
  STORE 0         TO MBWCASEPACK
  STORE .F.       TO MBWTRANSMIT
  STORE SPACE(13) TO MJITEMNO
  STORE 0         TO MJQTYAVAIL
  STORE 0         TO MJPRICE
  STORE 0         TO MJCASEPACK
  STORE .F.       TO MJTRANSMIT
  STORE SPACE(13) TO MbrITEMNO
  STORE 0         TO MbrQTYAVAIL
  STORE 0         TO MbrPRICE
  STORE 0         TO MbrCASEPACK
  STORE .F.       TO MbrTRANSMIT
  STORE SPACE(13) TO MBEITEMNO
  STORE 0         TO MBEQTYAVAIL
  STORE 0         TO MBEPRICE
  STORE 0         TO MBECASEPACK
  STORE .F.       TO MBETRANSMIT
  STORE SPACE(13) TO MWGITEMNO
  STORE 0         TO MWGQTYAVAIL
  STORE 0         TO MWGPRICE
  STORE 0         TO MWGCASEPACK
  STORE .F.       TO MWGVALID
  STORE .F.       TO MWGTRANSMIT
  STORE SPACE(13) TO MMCITEMNO
  STORE 0         TO MMCQTYAVAIL
  STORE 0         TO MMCPRICE
  STORE 0         TO MMCCASEPACK
  STORE "S"       TO MMCREF
  STORE .F.       TO MMCTRANSMIT
  STORE "EA"      TO MMCORDUNIT
  STORE SPACE(13) TO MCLITEMNO
  STORE 0         TO MCLQTYAVAIL
  STORE 0         TO MCLPRICE
  STORE 0         TO MCLCASEPACK
  STORE .F.       TO MCLAD
  STORE .F.       TO MCLTRANSMIT
  STORE .F.       TO MCLVALID
  STORE SPACE(13) TO MFDITEMNO
  STORE 0         TO MFDQTYAVAIL
  STORE 0         TO MFDPRICE
  STORE 0         TO MFDCASEPACK
  STORE .F.       TO MFDVALID
  STORE .F.       TO MFDTRANSMIT
  STORE SPACE(13) TO MHITEMNO
  STORE 0         TO MHQTYAVAIL
  STORE 0         TO MHPRICE
  STORE 0         TO MHCASEPACK
  STORE .F.       TO MHTRANSMIT
  STORE SPACE(13) TO MBDITEMNO
  STORE 0         TO MBDQTYAVAIL
  STORE 0         TO MBDPRICE
  STORE 0         TO MBDCASEPACK
  STORE .F.       TO MBDTRANSMIT
  STORE .T.       TO MBDBWUPDATE

  DO CASE
    CASE MCUSTNO=1
      @ 5,5 SAY PADR("Bindley Western",67) COLOR "GR+/N"
      @ 6,52 SAY "AD       = F3"
      FLD[1] = 'RITEM'
      FLD[2] = 'BWIN'
      FLD[3] = 'BWITEMNO'
      FLD[4] = 'BWQTYAVAIL'
      FLD[5] = 'BWPRICE'
      FLD[6] = 'BWCASEPACK'
      FLD[7] = 'CLAD'
      FLD[8] = 'BWTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'BW Item#'
      HEAD[3] = 'Item #'
      HEAD[4] = 'Qty;Avail'
      HEAD[5] = 'Price'
      HEAD[6] = 'Case;Pack'
      HEAD[7] = 'A;D'
      HEAD[8] = 'T;X'
      PIC[1] = 'XXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXX'
      PIC[3] = 'XXXXXXXXXXX'
      PIC[4] = '99999'
      PIC[5] = '9999.99'
      PIC[6] = '999'
      PIC[7] = 'Y'
      PIC[8] = 'Y'
    CASE MCUSTNO=2
      @ 5,5 SAY PADR("Kinray",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'FITEMNO'
      FLD[3] = 'FQTYAVAIL'
      FLD[4] = 'FPRICE'
      FLD[5] = 'FCASEPACK'
      FLD[6] = 'FTRANSMIT'
      FLD[7] = 'FVALID'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit?'
      HEAD[7] = 'Valid?'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = '999999'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
      PIC[7] = 'Y'
    CASE MCUSTNO=3
      @ 5,5 SAY PADR("H.D. Smith",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'TITEMNO'
      FLD[3] = 'TQTYAVAIL'
      FLD[4] = 'TPRICE'
      FLD[5] = 'TCASEPACK'
      FLD[6] = 'TTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
    CASE MCUSTNO=4
      @ 5,5 SAY PADR("Burlington Drug",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'WITEMNO'
      FLD[3] = 'WQTYAVAIL'
      FLD[4] = 'WPRICE'
      FLD[5] = 'WCASEPACK'
      FLD[6] = 'WTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
    CASE MCUSTNO=5
      @ 5,5 SAY PADR("Schein",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'NITEMNO'
      FLD[3] = 'NQTYAVAIL'
      FLD[4] = 'NPRICE'
      FLD[5] = 'NCASEPACK'
      FLD[6] = 'NTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
    CASE MCUSTNO=6
      @ 5,5 SAY PADR("Bergen-Brunswig",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'BITEMNO'
      FLD[3] = 'BQTYAVAIL'
      FLD[4] = 'BPRICE'
      FLD[5] = 'BCASEPACK'
      FLD[6] = 'BMINORDER'
      FLD[7] = 'BTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Min.;Order'
      HEAD[7] = 'Transmit?'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = '999999'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = '999'
      PIC[7] = 'Y'
    CASE MCUSTNO=7
      @ 5,5 SAY PADR("Rochester/Dakota",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'JITEMNO'
      FLD[3] = 'JQTYAVAIL'
      FLD[4] = 'JPRICE'
      FLD[5] = 'JCASEPACK'
      FLD[6] = 'JTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
//DGI(SG) 9/6/95 - Added this
    CASE MCUSTNO=8
      @ 5,5 SAY PADR("Walsh",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'BRITEMNO'
      FLD[3] = 'BRQTYAVAIL'
      FLD[4] = 'BRPRICE'
      FLD[5] = 'BRCASEPACK'
      FLD[6] = 'BRTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
    CASE MCUSTNO=9
      @ 5,5 SAY PADR("D & K Drug",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'BEITEMNO'
      FLD[3] = 'BEQTYAVAIL'
      FLD[4] = 'BEPRICE'
      FLD[5] = 'BECASEPACK'
      FLD[6] = 'BETRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
    CASE MCUSTNO=10
      @ 5,5 SAY PADR("Walgreens",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'WGITEMNO'
      FLD[3] = 'WGQTYAVAIL'
      FLD[4] = 'WGPRICE'
      FLD[5] = 'WGCASEPACK'
      FLD[6] = 'WGVALID'
      FLD[7] = 'WGTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'V'
      HEAD[7] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
      PIC[7] = 'Y'
    CASE MCUSTNO=11
      @ 5,5 SAY PADR("McKesson",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'MCITEMNO'
      FLD[3] = 'MCQTYAVAIL'
      FLD[4] = 'MCPRICE'
      FLD[5] = 'MCCASEPACK'
      FLD[6] = 'MCORDUNIT'
      FLD[7] = 'MCREFTYPE'
      FLD[8] = 'MCTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'OU'
      HEAD[7] = 'Typ'
      HEAD[8] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'XX'
      PIC[7] = 'X'
      PIC[8] = 'Y'
    CASE MCUSTNO=12
      @ 5,5 SAY PADR("Cardinal Health",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'CLITEMNO'
      FLD[3] = 'CLQTYAVAIL'
      FLD[4] = 'CLPRICE'
      FLD[5] = 'CLCASEPACK'
      FLD[6] = 'CLTRANSMIT'
      FLD[7] = 'CLVALID'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'TX'
      HEAD[7] = 'Valid'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
      PIC[7] = 'Y'
    CASE MCUSTNO=13
      @ 5,5 SAY PADR("F. Dohmen/Smith Drug",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'FDITEMNO'
      FLD[3] = 'FDQTYAVAIL'
      FLD[4] = 'FDPRICE'
      FLD[5] = 'FDCASEPACK'
      FLD[6] = 'FDVALID'
      FLD[7] = 'FDTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'V'
      HEAD[7] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
      PIC[7] = 'Y'
    CASE MCUSTNO=14
      @ 5,5 SAY PADR("Harvard Drug",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'HITEMNO'
      FLD[3] = 'HQTYAVAIL'
      FLD[4] = 'HPRICE'
      FLD[5] = 'HCASEPACK'
      FLD[6] = 'HTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Transmit'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = '@R99999-9999-99'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
    CASE MCUSTNO=15
      @ 5,5 SAY PADR("Bellco Drug/M&D",67) COLOR "GR+/N"
      FLD[1] = 'RITEM'
      FLD[2] = 'BDITEMNO'
      FLD[3] = 'BDQTYAVAIL'
      FLD[4] = 'BDPRICE'
      FLD[5] = 'BDCASEPACK'
      FLD[6] = 'BDBWUPDATE'
      FLD[7] = 'BDTRANSMIT'
      HEAD[1] = 'Item;Description'
      HEAD[2] = 'Item #'
      HEAD[3] = 'Qty;Avail'
      HEAD[4] = 'Price'
      HEAD[5] = 'Case;Pack'
      HEAD[6] = 'Bindley;Update'
      HEAD[7] = 'TX'
      PIC[1] = 'XXXXXXXXXXXXXXXXXX'
      PIC[2] = 'XXXXXXXXXXXXX'
      PIC[3] = '99999'
      PIC[4] = '9999.99'
      PIC[5] = '999'
      PIC[6] = 'Y'
      PIC[7] = 'Y'
      @ 6,52 SAY "BINDLEY UPDATE = F3"
  ENDCASE
  @ 5,52 SAY "TRANSMIT = F2"
  @ 7,52 SAY "FIND     = ANY CHARACTER"
  @ 8,52 SAY "CHANGE   = ENTER        "
  @ 9,52 SAY "EXIT     = ESC          "
  KEYBOARD CHR(19)
  DONE = .F.
  @ 10,3 TO 22,79 DOUBLE
  @ 11,4 CLEAR TO 21,78
  DO WHILE .NOT. DONE
    DBEDIT(11,4,21,78,FLD,"CUSTFUNC",PIC,HEAD)
  ENDDO
  RELEASE ALL
RETURN


FUNCTION CUSTFUNC
  PARAMETER STATUS, FLD_PTR

//(SG) 9/5/95 - Added Walsh

  RETVAL = 1
  MLAST = LASTKEY()
  DO CASE
    //====================== F2 - toggle transmit on / off
    CASE MLAST = -1
      IF RECLOCK(3)
        DO CASE
          CASE MCUSTNO = 1
            REPLACE BWTRANSMIT WITH !BWTRANSMIT
            REPLACE BTRANSMIT  WITH BWTRANSMIT
            REPLACE FTRANSMIT  WITH BWTRANSMIT
            REPLACE TTRANSMIT  WITH BWTRANSMIT
            REPLACE WTRANSMIT  WITH BWTRANSMIT
            REPLACE NTRANSMIT  WITH BWTRANSMIT
            REPLACE JTRANSMIT  WITH BWTRANSMIT
            REPLACE BRTRANSMIT WITH BWTRANSMIT
            REPLACE BETRANSMIT WITH BWTRANSMIT
            REPLACE WGTRANSMIT WITH BWTRANSMIT
            REPLACE CLTRANSMIT WITH BWTRANSMIT
            REPLACE MCTRANSMIT WITH BWTRANSMIT
            REPLACE FDTRANSMIT WITH BWTRANSMIT
            REPLACE HTRANSMIT  WITH BWTRANSMIT
            REPLACE BDTRANSMIT WITH BWTRANSMIT

          CASE MCUSTNO = 2
            REPLACE FTRANSMIT  WITH !FTRANSMIT
          CASE MCUSTNO = 3
            REPLACE TTRANSMIT  WITH !TTRANSMIT
          CASE MCUSTNO = 4
            REPLACE WTRANSMIT  WITH !WTRANSMIT
          CASE MCUSTNO = 5
            REPLACE NTRANSMIT  WITH !NTRANSMIT
          CASE MCUSTNO = 6
            REPLACE BTRANSMIT WITH !BTRANSMIT
          CASE MCUSTNO = 7
            REPLACE JTRANSMIT  WITH !JTRANSMIT
          CASE MCUSTNO = 8
            REPLACE BRTRANSMIT WITH !BRTRANSMIT
          CASE MCUSTNO = 9
            REPLACE BETRANSMIT WITH !BETRANSMIT
          CASE MCUSTNO = 10
            REPLACE WGTRANSMIT WITH !WGTRANSMIT
          CASE MCUSTNO = 11
            REPLACE MCTRANSMIT WITH !MCTRANSMIT
          CASE MCUSTNO = 12
            REPLACE CLTRANSMIT WITH !CLTRANSMIT
          CASE MCUSTNO = 13
            REPLACE FDTRANSMIT WITH !FDTRANSMIT
          CASE MCUSTNO = 14
            REPLACE HTRANSMIT WITH !HTRANSMIT
          CASE MCUSTNO = 15
            REPLACE BDTRANSMIT WITH !BDTRANSMIT
        ENDCASE
      ENDIF
      UNLOCK
      KEYBOARD CHR(24)
      RETVAL = 1

    CASE MLAST = -2  && F3
      IF RECLOCK(3)
        IF MCUSTNO = 1
          REPLACE CLAD WITH !CLAD
        ENDIF
        IF MCUSTNO = 15
          REPLACE BDBWUPDATE WITH !BDBWUPDATE
          IF !BDBWUPDATE
            REPLACE BDTRANSMIT WITH .F.
          ENDIF
        ENDIF
      ENDIF
      UNLOCK
      KEYBOARD CHR(24)
      RETVAL = 1

    CASE MLAST = 27                              && ESCAPE KEY IS PRESSED
      DONE  = .T.
      RETVAL = 0
    CASE MLAST = 13                              && ENTER KEY IS PRESSED
      SAVE SCREEN TO TEMP
      @ 6,52 CLEAR TO 9,75
      IF RECLOCK(3)
        DO CUSTSTORE
        DO CUSTGET WITH 'E'
        READ
        IF LASTKEY()<>27
          DO CUSTREPL
        ENDIF
        RESTORE SCREEN FROM TEMP
        UNLOCK
      ELSE
        MSGBOX("Record is in use.  Try again later.")
      ENDIF
      RETVAL = 1
    CASE (MLAST>=49.AND.MLAST<=122)
      SET CURSOR ON
      R = RECNO()
      MCHAR = CHR(MLAST)+SPACE(29)
      KEYBOARD CHR(4)
      @ 23,0 CLEAR
      @ 23,0 SAY "Find: " GET MCHAR PICTURE "@!"
      READ
      @ 23,0 CLEAR
      SEEK TRIM(MCHAR)
      IF FOUND()
        R=RECNO()
      ENDIF
      GOTO R
      KEYBOARD CHR(19)
      SET CURSOR OFF
      RETVAL = 0
  ENDCASE
RETURN RETVAL

PROCEDURE CUSTGET
PARAMETER MODE
//(SG) 9/5/95 - Added Walsh

  SET CURSOR ON
  SET CONFIRM ON
  @ 11, 20  TO 23, 66    DOUBLE
  @ 12, 21  CLEAR TO 22,65
  @ 12, 22  SAY "Enter Information:"
  DO CASE
    CASE MCUSTNO=1
      @ 14, 22  SAY "Item Desc  : "+MRITEM
      @ 15, 22  SAY "Bind Item #:" GET MBWIN PICT 'XXXXXX'
      @ 16, 22  SAY "Item #     :" GET MBWITEMNO PICT 'XXXXXXXXXXXXX'
      @ 17, 22  SAY "Qty Avail  :" GET MBWQTYAVAIL PICT '99999'
      @ 18, 22  SAY 'Price      :' GET MBWPRICE PICT '9999.99'
      @ 19, 22  SAY 'Case Pack  :' GET MBWCASEPACK PICT '999'
      @ 20, 22  SAY 'AD         :' GET MCLAD PICT 'Y'
      @ 21, 22  SAY 'Transmit?  :' GET MBWTRANSMIT PICT 'Y'
    CASE MCUSTNO=2
      @ 14, 22  SAY "Item Desc: "+MRITEM
      @ 15, 22  SAY "Item #   :" GET MFITEMNO PICT '999999'
      @ 16, 22  SAY "Qty Avail:" GET MFQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price    :' GET MFPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack:' GET MFCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit?:' GET MFTRANSMIT PICT 'Y'
      @ 20, 22  SAY 'Valid?   :' GET MFVALID    PICT 'Y'
    CASE MCUSTNO=3
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MTITEMNO PICT 'XXXXXXXXXXXXX'
      @ 16, 22  SAY "Qty Avail :" GET MTQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MTPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MTCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET MTTRANSMIT PICT 'Y'
    CASE MCUSTNO=4
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MWITEMNO PICT 'XXXXXXXXXXXXX'
      @ 16, 22  SAY "Qty Avail :" GET MWQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MWPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MWCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET MWTRANSMIT PICT 'Y'
    CASE MCUSTNO=5
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MNITEMNO PICT 'XXXXXXXXXXXXX'
      @ 16, 22  SAY "Qty Avail :" GET MNQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MNPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MNCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET MNTRANSMIT PICT 'Y'
    CASE MCUSTNO=6
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MBITEMNO PICT '999999'
      @ 16, 22  SAY "Qty Avail :" GET MBQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MBPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MBCASEPACK PICT '999'
      @ 19, 22  SAY 'Min. Order:' GET MBMINORDER PICT '999'
      @ 20, 22  SAY 'Transmit? :' GET MBTRANSMIT PICT 'Y'
    CASE MCUSTNO=7
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MJITEMNO PICT 'XXXXXXXXXXXXX'
      @ 16, 22  SAY "Qty Avail :" GET MJQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MJPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MJCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET MJTRANSMIT PICT 'Y'
    CASE MCUSTNO=8
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET mBRITEMNO PICT 'XXXXXXXXXXXXX'
      @ 16, 22  SAY "Qty Avail :" GET mBRQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET mBRPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET mBRCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET mBRTRANSMIT PICT 'Y'
    CASE MCUSTNO=9
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET mBEITEMNO PICT 'XXXXXXXXXXXXX' && VALID GETBELLCONUM(@mBEITEMNO)
      @ 16, 22  SAY "Qty Avail :" GET mBEQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET mBEPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET mBECASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET mBETRANSMIT PICT 'Y'
    CASE MCUSTNO=10
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MWGITEMNO PICT '999999'
      @ 16, 22  SAY "Qty Avail :" GET MWGQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MWGPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MWGCASEPACK PICT '999'
      @ 19, 22  SAY 'Validated :' GET MWGVALID PICT 'Y'
      @ 20, 22  SAY 'Transmit? :' GET MWGTRANSMIT PICT 'Y'
    CASE MCUSTNO=11
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MMCITEMNO PICT '9999999'
      @ 16, 22  SAY "Qty Avail :" GET MMCQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MMCPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MMCCASEPACK PICT '999'
      @ 19, 22  SAY 'Order Unit:' GET MMCORDUNIT PICT '!!' VALID CHECKORDUNIT(@MMCORDUNIT)
      @ 20, 22  SAY 'Type      :' GET MMCREF PICT '!' VALID MMCREF $ "SIR" 
      * CHECKREFTYPE(@MMCREF)
      @ 20, 36  SAY '(R=Refrig,I=Inject,S=Standard)'
      @ 21, 22  SAY 'Transmit? :' GET MMCTRANSMIT PICT 'Y'
    CASE MCUSTNO=12
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MCLITEMNO PICT '9999999'
      @ 16, 22  SAY "Qty Avail :" GET MCLQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MCLPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MCLCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET MCLTRANSMIT PICT 'Y'
      @ 20, 22  SAY 'Valid?    :' GET MCLVALID    PICT 'Y'
    CASE MCUSTNO=13
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MFDITEMNO PICT '99999999999'
      @ 16, 22  SAY "Qty Avail :" GET MFDQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MFDPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MFDCASEPACK PICT '999'
      @ 19, 22  SAY 'Validated :' GET MFDVALID PICT 'Y'
      @ 20, 22  SAY 'Transmit? :' GET MFDTRANSMIT PICT 'Y'
    CASE MCUSTNO=14
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MHITEMNO PICT '99999-9999-99'
      @ 16, 22  SAY "Qty Avail :" GET MHQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MHPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MHCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET MHTRANSMIT PICT 'Y'
    CASE MCUSTNO=15
      @ 14, 22  SAY "Item Desc : "+MRITEM
      @ 15, 22  SAY "Item #    :" GET MBDITEMNO PICT 'XXXXXXXXXXXXX'
      @ 16, 22  SAY "Qty Avail :" GET MBDQTYAVAIL PICT '99999'
      @ 17, 22  SAY 'Price     :' GET MBDPRICE PICT '9999.99'
      @ 18, 22  SAY 'Case Pack :' GET MBDCASEPACK PICT '999'
      @ 19, 22  SAY 'Transmit? :' GET MBDTRANSMIT PICT 'Y'
      @ 20, 22  SAY 'Update From Bindley? :' GET MBDBWUPDATE PICT 'Y'
  ENDCASE
RETURN

PROCEDURE CUSTINIT
//(SG) 9/5/95 - Added Walsh
  STORE space(30) TO MRITEM
  STORE space(6)  TO MFITEMNO
  STORE 0         TO MFPRICE
  STORE 0         TO MFQTYAVAIL
  STORE 0         TO MFCASEPACK
  STORE .F.       TO MFTRANSMIT
  STORE .F.       TO MFVALID
  STORE 0         TO MBITEMNO
  STORE 0         TO MBQTYAVAIL
  STORE 0         TO MBPRICE
  STORE 0         TO MBCASEPACK
  STORE 0         TO MBMINORDER
  STORE .F.       TO MBTRANSMIT
  STORE SPACE(13) TO MTITEMNO
  STORE 0         TO MTQTYAVAIL
  STORE 0         TO MTPRICE
  STORE 0         TO MTCASEPACK
  STORE .F.       TO MTTRANSMIT
  STORE SPACE(13) TO MWITEMNO
  STORE 0         TO MWQTYAVAIL
  STORE 0         TO MWPRICE
  STORE 0         TO MWCASEPACK
  STORE .F.       TO MWTRANSMIT
  STORE SPACE(13) TO MNITEMNO
  STORE 0         TO MNQTYAVAIL
  STORE 0         TO MNPRICE
  STORE 0         TO MNCASEPACK
  STORE .F.       TO MNTRANSMIT
  STORE SPACE(6)  TO MBWIN
  STORE SPACE(13) TO MBWITEMNO
  STORE 0         TO MBWQTYAVAIL
  STORE 0         TO MBWPRICE
  STORE 0         TO MBWCASEPACK
  STORE .F.       TO MBWTRANSMIT
  STORE SPACE(13) TO MJITEMNO
  STORE 0         TO MJQTYAVAIL
  STORE 0         TO MJPRICE
  STORE 0         TO MJCASEPACK
  STORE .F.       TO MJTRANSMIT

  STORE SPACE(13) TO MBRITEMNO
  STORE 0         TO MBRQTYAVAIL
  STORE 0         TO MBRPRICE
  STORE 0         TO MBRCASEPACK
  STORE .F.       TO MBRTRANSMIT

  STORE SPACE(13) TO MBEITEMNO
  STORE 0         TO MBEQTYAVAIL
  STORE 0         TO MBEPRICE
  STORE 0         TO MBECASEPACK
  STORE .F.       TO MBETRANSMIT

  STORE SPACE(13) TO MWGITEMNO
  STORE 0         TO MWGQTYAVAIL
  STORE 0         TO MWGPRICE
  STORE 0         TO MWGCASEPACK
  STORE .F.       TO MWGVALID
  STORE .F.       TO MWGTRANSMIT

  STORE SPACE(13) TO MMCITEMNO
  STORE 0         TO MMCQTYAVAIL
  STORE 0         TO MMCPRICE
  STORE 0         TO MMCCASEPACK
  STORE "EA"      TO MMCORDUNIT
  STORE "S"       TO MMCREF
  STORE .F.       TO MMCTRANSMIT

  STORE SPACE(13) TO MCLITEMNO
  STORE 0         TO MCLQTYAVAIL
  STORE 0         TO MCLPRICE
  STORE 0         TO MCLCASEPACK
  STORE .F.       TO MCLTRANSMIT
  STORE .F.       TO MCLAD
  STORE .F.       TO MCLVALID

  STORE SPACE(13) TO MFDITEMNO
  STORE 0         TO MFDQTYAVAIL
  STORE 0         TO MFDPRICE
  STORE 0         TO MFDCASEPACK
  STORE .F.       TO MFDVALID
  STORE .F.       TO MFDTRANSMIT

  STORE SPACE(13) TO MHITEMNO
  STORE 0         TO MHQTYAVAIL
  STORE 0         TO MHPRICE
  STORE 0         TO MHCASEPACK
  STORE .F.       TO MHTRANSMIT

  STORE SPACE(13) TO MBDITEMNO
  STORE 0         TO MBDQTYAVAIL
  STORE 0         TO MBDPRICE
  STORE 0         TO MBDCASEPACK
  STORE .F.       TO MBDTRANSMIT
  STORE .T.       TO MBDBWUPDATE
RETURN

PROCEDURE CUSTSTORE
  STORE RITEM TO MRITEM

  STORE BITEMNO TO MBITEMNO
  STORE BQTYAVAIL TO MBQTYAVAIL
  STORE BPRICE TO MBPRICE
  STORE BCASEPACK TO MBCASEPACK
  STORE BMINORDER TO MBMINORDER
  STORE BTRANSMIT TO MBTRANSMIT

  STORE FITEMNO TO MFITEMNO
  STORE FCASEPACK TO MFCASEPACK
  STORE FPRICE TO MFPRICE
  STORE FQTYAVAIL TO MFQTYAVAIL
  STORE FTRANSMIT TO MFTRANSMIT
  STORE FVALID TO MFVALID

  STORE TITEMNO TO MTITEMNO
  STORE TCASEPACK TO MTCASEPACK
  STORE TQTYAVAIL TO MTQTYAVAIL
  STORE TPRICE TO MTPRICE
  STORE TTRANSMIT TO MTTRANSMIT

  STORE WITEMNO   TO MWITEMNO
  STORE WCASEPACK TO MWCASEPACK
  STORE WQTYAVAIL TO MWQTYAVAIL
  STORE WPRICE    TO MWPRICE
  STORE WTRANSMIT TO MWTRANSMIT

  STORE NITEMNO   TO MNITEMNO
  STORE NCASEPACK TO MNCASEPACK
  STORE NQTYAVAIL TO MNQTYAVAIL
  STORE NPRICE    TO MNPRICE
  STORE NTRANSMIT TO MNTRANSMIT

  STORE BWIN       TO MBWIN
  STORE BWITEMNO   TO MBWITEMNO
  STORE BWCASEPACK TO MBWCASEPACK
  STORE BWQTYAVAIL TO MBWQTYAVAIL
  STORE BWPRICE    TO MBWPRICE
  STORE BWTRANSMIT TO MBWTRANSMIT

  STORE JITEMNO   TO MJITEMNO
  STORE JCASEPACK TO MJCASEPACK
  STORE JQTYAVAIL TO MJQTYAVAIL
  STORE JPRICE    TO MJPRICE
  STORE JTRANSMIT TO MJTRANSMIT

  STORE BRITEMNO   TO MBRITEMNO
  STORE BRCASEPACK TO MBRCASEPACK
  STORE BRQTYAVAIL TO MBRQTYAVAIL
  STORE BRPRICE    TO MBRPRICE
  STORE BRTRANSMIT TO MBRTRANSMIT

  STORE BEITEMNO   TO MBEITEMNO
  STORE BECASEPACK TO MBECASEPACK
  STORE BEQTYAVAIL TO MBEQTYAVAIL
  STORE BEPRICE    TO MBEPRICE
  STORE BETRANSMIT TO MBETRANSMIT

  STORE WGITEMNO   TO MWGITEMNO
  STORE WGCASEPACK TO MWGCASEPACK
  STORE WGQTYAVAIL TO MWGQTYAVAIL
  STORE WGPRICE    TO MWGPRICE
  STORE WGVALID    TO MWGVALID
  STORE WGTRANSMIT TO MWGTRANSMIT

  STORE MCITEMNO   TO MMCITEMNO
  STORE MCCASEPACK TO MMCCASEPACK
  STORE MCQTYAVAIL TO MMCQTYAVAIL
  STORE MCPRICE    TO MMCPRICE
  IF EMPTY(MCORDUNIT)
    STORE "EA" TO MMORDUNIT
  ELSE
    STORE MCORDUNIT  TO MMCORDUNIT
  ENDIF
  STORE MCREFTYPE TO MMCREF
  STORE MCTRANSMIT TO MMCTRANSMIT

  STORE CLITEMNO   TO MCLITEMNO
  STORE CLCASEPACK TO MCLCASEPACK
  STORE CLQTYAVAIL TO MCLQTYAVAIL
  STORE CLPRICE    TO MCLPRICE
  STORE CLTRANSMIT TO MCLTRANSMIT
  STORE CLAD       TO MCLAD
  STORE CLVALID    TO MCLVALID

  STORE FDITEMNO   TO MFDITEMNO
  STORE FDCASEPACK TO MFDCASEPACK
  STORE FDQTYAVAIL TO MFDQTYAVAIL
  STORE FDPRICE    TO MFDPRICE
  STORE FDVALID    TO MFDVALID
  STORE FDTRANSMIT TO MFDTRANSMIT

  STORE HITEMNO   TO MHITEMNO
  STORE HCASEPACK TO MHCASEPACK
  STORE HQTYAVAIL TO MHQTYAVAIL
  STORE HPRICE    TO MHPRICE
  STORE HTRANSMIT TO MHTRANSMIT

  STORE BDITEMNO   TO MBDITEMNO
  STORE BDCASEPACK TO MBDCASEPACK
  STORE BDQTYAVAIL TO MBDQTYAVAIL
  STORE BDPRICE    TO MBDPRICE
  STORE BDTRANSMIT TO MBDTRANSMIT
  STORE BDBWUPDATE TO MBDBWUPDATE
RETURN


PROCEDURE CUSTREPL
    REPLACE RITEM WITH MRITEM

    REPLACE BWQTYAVAIL WITH MBWQTYAVAIL
    REPLACE BWIN       WITH MBWIN
    REPLACE BWITEMNO   WITH MBWITEMNO
    REPLACE BWPRICE    WITH MBWPRICE
    REPLACE BWCASEPACK WITH MBWCASEPACK
    REPLACE BWTRANSMIT WITH MBWTRANSMIT
    REPLACE CLAD        WITH MCLAD

    REPLACE BMINORDER WITH MBMINORDER

    // SIMON
    // AUTO-UPDATE OF QTY,PRICE,TRANS WHEN BINDLEY IS ENTERED.
    IF MCUSTNO == 1  && BINDLEY
      REPLACE BWCASEPACK WITH MBWCASEPACK
      REPLACE NCASEPACK  WITH MBWCASEPACK
      REPLACE TCASEPACK  WITH MBWCASEPACK
      REPLACE JCASEPACK  WITH MBWCASEPACK
      REPLACE BRCASEPACK WITH MBWCASEPACK
      REPLACE BECASEPACK WITH MBWCASEPACK
      REPLACE WGCASEPACK WITH MBWCASEPACK
      REPLACE CLCASEPACK WITH MBWCASEPACK
      REPLACE FDCASEPACK WITH MBWCASEPACK
      REPLACE HCASEPACK  WITH MBWCASEPACK
      REPLACE FCASEPACK  WITH MBWCASEPACK
      REPLACE WCASEPACK  WITH MBWCASEPACK
      REPLACE BCASEPACK  WITH MBWCASEPACK
      REPLACE MCCASEPACK WITH MBWCASEPACK

      // REPLACE NITEMNO    WITH MBWITEMNO
      // REPLACE WITEMNO    WITH MBWITEMNO
      // REPLACE BEITEMNO   WITH MBWITEMNO
      // DON'T REPLACE IF H.D. Smith IS ALREADY AN 11 DIGIT NUMBER
      // IF LEN(TRIM(TITEMNO)) < 11
        // REPLACE TITEMNO    WITH MBWITEMNO
      // ENDIF
      // REPLACE JITEMNO    WITH MBWITEMNO
      // REPLACE BRITEMNO   WITH MBWITEMNO
      // REPLACE FDITEMNO   WITH MBWITEMNO

      REPLACE FQTYAVAIL WITH MBWQTYAVAIL
      REPLACE FTRANSMIT WITH MBWTRANSMIT
      REPLACE FPRICE    WITH MBWPRICE

      REPLACE TQTYAVAIL WITH MBWQTYAVAIL
      REPLACE TPRICE    WITH MBWPRICE
      REPLACE TTRANSMIT WITH MBWTRANSMIT

      REPLACE WQTYAVAIL WITH MBWQTYAVAIL
      REPLACE WPRICE    WITH MBWPRICE
      REPLACE WTRANSMIT WITH MBWTRANSMIT

      REPLACE NQTYAVAIL WITH MBWQTYAVAIL
      REPLACE NPRICE    WITH MBWPRICE
      REPLACE NTRANSMIT WITH MBWTRANSMIT

      REPLACE JQTYAVAIL WITH MBWQTYAVAIL
      REPLACE JPRICE    WITH MBWPRICE
      REPLACE JTRANSMIT WITH MBWTRANSMIT

      REPLACE BRQTYAVAIL WITH MBWQTYAVAIL
      REPLACE BRPRICE    WITH MBWPRICE
      REPLACE BRTRANSMIT WITH MBWTRANSMIT

      REPLACE BEQTYAVAIL WITH MBWQTYAVAIL
      REPLACE BEPRICE    WITH MBWPRICE
      REPLACE BETRANSMIT WITH MBWTRANSMIT

      REPLACE BQTYAVAIL  WITH MBWQTYAVAIL
      REPLACE BPRICE     WITH MBWPRICE
      *REPLACE BCASEPACK  WITH MBWCASEPACK

      REPLACE WGQTYAVAIL WITH MBWQTYAVAIL
      REPLACE WGTRANSMIT WITH MBWTRANSMIT
      REPLACE WGPRICE    WITH MBWPRICE

      REPLACE CLQTYAVAIL  WITH MBWQTYAVAIL
      REPLACE CLPRICE     WITH MBWPRICE
      REPLACE CLCASEPACK  WITH MBWCASEPACK

      REPLACE MCQTYAVAIL WITH MBWQTYAVAIL
      REPLACE MCPRICE    WITH MBWPRICE
      REPLACE MCTRANSMIT WITH MBWTRANSMIT

      REPLACE FDQTYAVAIL WITH MBWQTYAVAIL
      REPLACE FDTRANSMIT WITH MBWTRANSMIT
      REPLACE FDPRICE    WITH MBWPRICE

      REPLACE HQTYAVAIL WITH MBWQTYAVAIL
      REPLACE HTRANSMIT WITH MBWTRANSMIT
      REPLACE HPRICE    WITH MBWPRICE

      REPLACE BDQTYAVAIL WITH MBWQTYAVAIL
      REPLACE BDPRICE    WITH MBWPRICE
      IF MBDBWUPDATE
        REPLACE BDTRANSMIT WITH MBWTRANSMIT
      ENDIF
    ELSE
      REPLACE BCASEPACK  WITH MBCASEPACK
      REPLACE NCASEPACK  WITH MNCASEPACK
      REPLACE TCASEPACK  WITH MTCASEPACK
      REPLACE JCASEPACK  WITH MJCASEPACK
      REPLACE BRCASEPACK WITH MBRCASEPACK
      REPLACE BECASEPACK WITH MBECASEPACK
      REPLACE WGCASEPACK WITH MWGCASEPACK
      REPLACE MCCASEPACK WITH MMCCASEPACK
      REPLACE FDCASEPACK WITH MFDCASEPACK
      REPLACE HCASEPACK  WITH MHCASEPACK
      REPLACE FCASEPACK  WITH MFCASEPACK
      REPLACE WCASEPACK  WITH MWCASEPACK
      REPLACE BDCASEPACK WITH MBDCASEPACK

      REPLACE TITEMNO    WITH MTITEMNO
      REPLACE NITEMNO    WITH MNITEMNO
      REPLACE JITEMNO    WITH MJITEMNO
      REPLACE BRITEMNO   WITH MBRITEMNO
      REPLACE WITEMNO    WITH MWITEMNO
      REPLACE MCITEMNO   WITH MMCITEMNO
      REPLACE CLITEMNO   WITH MCLITEMNO
      REPLACE BEITEMNO   WITH MBEITEMNO
      REPLACE FDITEMNO   WITH MFDITEMNO
      REPLACE HITEMNO    WITH MHITEMNO
      REPLACE BDITEMNO   WITH MBDITEMNO

      REPLACE FQTYAVAIL WITH MFQTYAVAIL
      REPLACE FTRANSMIT WITH MFTRANSMIT
      REPLACE FPRICE    WITH MFPRICE
      REPLACE FVALID    WITH MFVALID

      REPLACE TQTYAVAIL WITH MTQTYAVAIL
      REPLACE TPRICE    WITH MTPRICE
      REPLACE TTRANSMIT WITH MTTRANSMIT

      REPLACE WQTYAVAIL WITH MWQTYAVAIL
      REPLACE WPRICE    WITH MWPRICE
      REPLACE WTRANSMIT WITH MWTRANSMIT

      REPLACE NQTYAVAIL WITH MNQTYAVAIL
      REPLACE NPRICE    WITH MNPRICE
      REPLACE NTRANSMIT WITH MNTRANSMIT

      REPLACE BTRANSMIT WITH MBTRANSMIT

      REPLACE JQTYAVAIL WITH MJQTYAVAIL
      REPLACE JPRICE    WITH MJPRICE
      REPLACE JTRANSMIT WITH MJTRANSMIT

      REPLACE BRQTYAVAIL WITH MBRQTYAVAIL
      REPLACE BRPRICE    WITH MBRPRICE
      REPLACE BRTRANSMIT WITH MBRTRANSMIT

      REPLACE BEQTYAVAIL WITH MBEQTYAVAIL
      REPLACE BEPRICE    WITH MBEPRICE
      REPLACE BETRANSMIT WITH MBETRANSMIT

      REPLACE BQTYAVAIL  WITH MBQTYAVAIL
      REPLACE BPRICE     WITH MBPRICE
      REPLACE BCASEPACK  WITH MBCASEPACK

      REPLACE WGQTYAVAIL WITH MWGQTYAVAIL
      REPLACE WGTRANSMIT WITH MWGTRANSMIT
      REPLACE WGVALID    WITH MWGVALID
      REPLACE WGPRICE    WITH MWGPRICE

      REPLACE MCQTYAVAIL WITH MMCQTYAVAIL
      REPLACE MCTRANSMIT WITH MMCTRANSMIT
      REPLACE MCREFTYPE  WITH MMCREF
      REPLACE MCORDUNIT  WITH MMCORDUNIT
      REPLACE MCPRICE    WITH MMCPRICE

      REPLACE CLQTYAVAIL  WITH MCLQTYAVAIL
      REPLACE CLPRICE     WITH MCLPRICE
      REPLACE CLCASEPACK  WITH MCLCASEPACK
      REPLACE CLTRANSMIT  WITH MCLTRANSMIT
      REPLACE CLVALID     WITH MCLVALID

      REPLACE FDQTYAVAIL WITH MFDQTYAVAIL
      REPLACE FDTRANSMIT WITH MFDTRANSMIT
      REPLACE FDVALID    WITH MFDVALID
      REPLACE FDPRICE    WITH MFDPRICE

      REPLACE HQTYAVAIL WITH MHQTYAVAIL
      REPLACE HTRANSMIT WITH MHTRANSMIT
      REPLACE HPRICE    WITH MHPRICE

      REPLACE BDQTYAVAIL WITH MBDQTYAVAIL
      REPLACE BDPRICE    WITH MBDPRICE
      REPLACE BDTRANSMIT WITH MBDTRANSMIT
      REPLACE BDBWUPDATE WITH MBDBWUPDATE
    ENDIF

    REPLACE FITEMNO   WITH MFITEMNO
    REPLACE BITEMNO   WITH MBITEMNO
    REPLACE WGITEMNO  WITH MWGITEMNO
RETURN


PROCEDURE RDREPT
  LOCAL MCUST[16]
  LOCAL MREPT[5]
  DO WHILE .T.

    MMFGFROM := "   "
    MMFGTO   := "   "

    MREPT[1]='Pre-Transmission Report'
    MREPT[2]='Master Inventory List'
    MREPT[3]='Bergen Purchase Order'
    MREPT[4]='McKesson Purchase Order'
    MREPT[5]='Return to Main Menu'
    MREPTNO=1
    MCUSTNO=1

    SELECT RDINV
    SET ORDER TO TAG RDINV  &&,RDBITEM
    GO TOP

    @ 5,5 SAY 'Select Report to Print...'
    @ 8,6 CLEAR TO 22,44
    @ 15,47 SAY "ENTER = SELECT"
    @ 16,47 SAY "ESC   = QUIT  "
    MREPTNO=ACHOICE(10,6,16,44,MREPT,.T.,'',MREPTNO)
    IF .NOT.(LASTKEY()=27.OR.MREPTNO=5)

      DO CASE
        CASE MREPTNO = 1
          * MCUST[1]='Bindley Western'
          * MCUST[2]='Kinray'
          * MCUST[3]='H.D. Smith'
          * MCUST[4]='Burlington/Amerisource Drug'
          * MCUST[5]='Schein'
          * MCUST[6]='Bergen-Brunswig'
          * MCUST[7]='Rochester/Dakota'
          * MCUST[8]='Walsh'
          * MCUST[9]='D & K Drug'
          * MCUST[10]='Walgreens'
          * MCUST[11]='McKesson'
          * MCUST[12]='Cardinal Health'
          * MCUST[13]='F. Dohmen/Smith Drug'
          * MCUST[14]='Harvard Drug'
          * MCUST[15]='Bellco Drug/M&D'
			MCUST[1]='Bindley Western'
			MCUST[2]='Walsh'
			MCUST[3]='Bellco Drug/M&D'
			MCUST[4]='Bergen-Brunswig'
			MCUST[5]='Burlington/Amerisource/NC Mutual/Longs'
			MCUST[6]='Cardinal Health'
			MCUST[7]='D & K Drug'
			MCUST[8]='F. Dohmen/Smith Drug'
			MCUST[9]='H.D. Smith'
			MCUST[10]='Harvard Drug'
			MCUST[11]='Kinray'
			MCUST[12]='McKesson'
			MCUST[13]='Schein'
			MCUST[14]='Rochester/Dakota'
			MCUST[15]='Walgreens'
          MCUST[16]='Return to Main Menu'

          @ 5,5 SAY 'Select Customer to Report...'
          @ 7,6 CLEAR TO 22,44
          @ 15,47 SAY "ENTER = SELECT"
          @ 16,47 SAY "ESC   = QUIT  "
          MCUSTNO=ACHOICE(7,6,22,44,MCUST,.T.,'',MCUSTNO)

		  ORGCUSTNO = MCUSTNO
		  	
			DO CASE
				CASE MCUSTNO = 1
					MCUSTNO = 1
				CASE MCUSTNO = 2
					MCUSTNO = 8
				CASE MCUSTNO = 3
					MCUSTNO = 15
				CASE MCUSTNO = 4
					MCUSTNO = 6
				CASE MCUSTNO = 5
					MCUSTNO = 4
				CASE MCUSTNO = 6
					MCUSTNO = 12
				CASE MCUSTNO = 7
					MCUSTNO = 9
				CASE MCUSTNO = 8
					MCUSTNO = 13
				CASE MCUSTNO = 9
					MCUSTNO = 3
				CASE MCUSTNO = 10
					MCUSTNO = 14
				CASE MCUSTNO = 11
					MCUSTNO = 2
				CASE MCUSTNO = 12
					MCUSTNO = 11
				CASE MCUSTNO = 13
					MCUSTNO = 5
				CASE MCUSTNO = 14
					MCUSTNO = 7
				CASE MCUSTNO = 15
					MCUSTNO = 10
			ENDCASE
          IF .NOT.(LASTKEY()=27.OR.MCUSTNO=16)
            GETMFGRANGE()
          ELSE
            @ 7,6 CLEAR TO 22,44
          ENDIF

          IF mCUSTNO=11
            mREF="S"
            @23,0 CLEAR
            @23,0 SAY "(S)tandard, (R)efrigerated, or (I)njectable? "
            @23,45 GET mREF PICTURE "@!" VALID mREF $ "SRI"
            READ
            @23,0 CLEAR
          ENDIF

			IF MCUSTNO=4
			  SSCREEN=SAVESCREEN(6,14,17,65)
			  @ 6,14 CLEAR TO 17,65
			  @ 7,15,16,64 BOX B_SINGLE+" "
			  @ 11,22 SAY "Please select a company to print:"
			  @ 13,18 PROMPT "Amerisource"
			  @ 13,31 PROMPT "Burlington"
			  @ 13,43 PROMPT "NC Mutual"
			  @ 13,54 PROMPT "Longs"
			  MENU TO MENUTO
			  RESTSCREEN(6,14,17,65,SSCREEN)
			ENDIF
          IF MCUSTNO=12 .OR. MCUSTNO=11 .OR. MCUSTNO=6
			  SSCREEN=SAVESCREEN(6,14,17,65)
			  @ 6,14 CLEAR TO 17,65
			  @ 7,15,16,64 BOX B_SINGLE+" "
			  @ 11,22 SAY "Please select a report type:"
			  @ 13,25 PROMPT "    (AD)      "
			  @ 13,41 PROMPT "    (All)     "
			  MENU TO MENUTO
			  RESTSCREEN(6,14,17,65,SSCREEN)
          ENDIF

          IF .NOT.(LASTKEY()=27.OR.MCUSTNO=16)
            MYORN=.F.
            @ 23,0 SAY 'Is the printer ready? [y/n]' GET MYORN PICT 'Y'
            READ
            IF MYORN
              * @ 23,0 CLEAR
              * @ 23,0 SAY 'Printing.  Please wait...'
              SELECT RDINV
              SET ORDER TO TAG RDINV  &&,RDBITEM
              GO TOP

              STATUSBOX("Printing "+MCUST[ORGCUSTNO]+" Inventory...")

              DO CASE
                CASE MCUSTNO=1
                  * @ 6,5 SAY "Bindley Western"
                  REPORT FORM RDBW FOR (BWTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=2
                  * @ 6,5 SAY "Kinray"
                  REPORT FORM RDKIN FOR (FTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=3
                  * @ 6,5 SAY "H.D. Smith"
                  REPORT FORM RDTEXAS FOR (TTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=4
                  do case
					case MENUTO = 1
	                  REPORT FORM RDAM FOR (WTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
					case MENUTO = 2
	                  REPORT FORM RDWALK FOR (WTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
	                case MENUTO = 3
	                  REPORT FORM RDNC FOR (WTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
	                case MENUTO = 4
	                  REPORT FORM RDLONGS FOR (WTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
	              endcase
                CASE MCUSTNO=5
                  * @ 6,5 SAY "Schein"
                  REPORT FORM RDNEU FOR (NTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=6
                  * @ 6,5 SAY "Bergen-Brunswig"
                  IF MENUTO==1
	                  REPORT FORM RDBERG FOR (BTRANSMIT.AND.CLAD.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                  ELSE
	                  REPORT FORM RDBERG FOR (BTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                  ENDIF
                CASE MCUSTNO=7
                  * @ 6,5 SAY "Rochester/Dakota"
                  REPORT FORM RDBN FOR (JTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=8
                  * @ 6,5 SAY "Walsh"
                  REPORT FORM RDBR FOR (BRTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=9
                  * @ 6,5 SAY "D & K Drug"
                  REPORT FORM RDDK FOR (BETRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=10
                  * @ 6,5 SAY "Walgreens"
                  REPORT FORM RDWG FOR (WGTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=11
                  * @ 6,5 SAY "McKesson"
                  * IF mREF="S"
                  
                  IF MENUTO==1
                    REPORT FORM RDMC FOR (MCREFTYPE==mREF.AND.MCTRANSMIT.AND.CLAD.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                  ELSE
                    REPORT FORM RDMC FOR (MCREFTYPE==mREF.AND.MCTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                  ENDIF
                  * ELSE
				  * 	REPORT FORM RDMC FOR (MCREF="I").AND.MCTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                  * ENDIF
                CASE MCUSTNO=12
                  * @ 6,5 SAY "Cardinal Health"
                  IF MENUTO==1
                    REPORT FORM RDCL FOR (CLTRANSMIT.AND.CLAD.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                  ELSE
                    REPORT FORM RDCL FOR (CLTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
				  ENDIF
                CASE MCUSTNO=13
                  * @ 6,5 SAY "F. Dohmen/Smith Drug"
                  REPORT FORM RDFD FOR (FDTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=14
                  * @ 6,5 SAY "Harvard Drug"
                  REPORT FORM RDH FOR (HTRANSMIT.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE MCUSTNO=15
                  * @ 6,5 SAY "Bellco Drug/M&D"
                  REPORT FORM RDBD FOR (BDTRANSMIT.AND.BDBWUPDATE.AND.WITHINRANGE()) TO PRINTER NOCONSOLE
                CASE LASTKEY()=27.OR.MCUSTNO=16
              ENDCASE

              CLEARSTATUS()

              * @ 23,0 CLEAR
              * @ 23,0 SAY "Report Printing Completed.  Press Any Key..."
              *INKEY(0)
              * @ 5,5 SAY SPACE(67)

              MSGBOX("Report Printing Completed.  Press Any Key...")

            ENDIF
          ENDIF
 //-------------------------------------- Stop Here ---------
        CASE MREPTNO = 2

// start
			MCUST[1]='Bindley Western'
			MCUST[2]='Walsh'
			MCUST[3]='Bellco Drug/M&D'
			MCUST[4]='Bergen-Brunswig'
			MCUST[5]='Burlington/Amerisource/NC Mutual/Longs'
			MCUST[6]='Cardinal Health'
			MCUST[7]='D & K Drug'
			MCUST[8]='F. Dohmen/Smith Drug'
			MCUST[9]='H.D. Smith'
			MCUST[10]='Harvard Drug'
			MCUST[11]='Kinray'
			MCUST[12]='McKesson'
			MCUST[13]='Schein'
			MCUST[14]='Rochester/Dakota'
			MCUST[15]='Walgreens'
          MCUST[16]='Return to Main Menu'

          @ 5,5 SAY 'Select Customer to Report...'
          @ 7,6 CLEAR TO 22,44
          @ 15,47 SAY "ENTER = SELECT"
          @ 16,47 SAY "ESC   = QUIT  "
          MCUSTNO=ACHOICE(7,6,22,44,MCUST,.T.,'',MCUSTNO)

		  ORGCUSTNO = MCUSTNO
		  	
			DO CASE
				CASE MCUSTNO = 1
					MCUSTNO = 1
				CASE MCUSTNO = 2
					MCUSTNO = 8
				CASE MCUSTNO = 3
					MCUSTNO = 15
				CASE MCUSTNO = 4
					MCUSTNO = 6
				CASE MCUSTNO = 5
					MCUSTNO = 4
				CASE MCUSTNO = 6
					MCUSTNO = 12
				CASE MCUSTNO = 7
					MCUSTNO = 9
				CASE MCUSTNO = 8
					MCUSTNO = 13
				CASE MCUSTNO = 9
					MCUSTNO = 3
				CASE MCUSTNO = 10
					MCUSTNO = 14
				CASE MCUSTNO = 11
					MCUSTNO = 2
				CASE MCUSTNO = 12
					MCUSTNO = 11
				CASE MCUSTNO = 13
					MCUSTNO = 5
				CASE MCUSTNO = 14
					MCUSTNO = 7
				CASE MCUSTNO = 15
					MCUSTNO = 10
			ENDCASE
          IF .NOT.(LASTKEY()=27.OR.MCUSTNO=16)
            GETMFGRANGE()
          ELSE
            @ 7,6 CLEAR TO 22,44
          ENDIF

          IF .NOT.(LASTKEY()=27.OR.MCUSTNO=16)
            MYORN=.F.
            @ 23,0 SAY 'Is the printer ready? [y/n]' GET MYORN PICT 'Y'
            READ
            IF MYORN
              * @ 23,0 CLEAR
              * @ 23,0 SAY 'Printing.  Please wait...'
              SELECT RDINV
              SET ORDER TO TAG RDINV  &&,RDBITEM
              GO TOP

              STATUSBOX("Printing "+MCUST[ORGCUSTNO]+" Inventory...")


              DO CASE
                CASE MCUSTNO=1
                  * @ 6,5 SAY "Bindley Western"
                  REPORT FORM RDMBW TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=2
                  * @ 6,5 SAY "Kinray"
                  REPORT FORM RDMKIN TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=3
                  * @ 6,5 SAY "H.D. Smith"
                  REPORT FORM RDMHD TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=4
                  REPORT FORM RDMAWNL TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=5
                  * @ 6,5 SAY "Schein"
                  REPORT FORM RDMNEU TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=6
                  * @ 6,5 SAY "Bergen-Brunswig"
                  REPORT FORM RDMBERG TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=7
                  * @ 6,5 SAY "Rochester/Dakota"
                  REPORT FORM RDMBN TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=8
                  * @ 6,5 SAY "Walsh"
                  REPORT FORM RDMBR TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=9
                  * @ 6,5 SAY "D & K Drug"
                  REPORT FORM RDMDK TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=10
                  * @ 6,5 SAY "Walgreens"
                  REPORT FORM RDMWG TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=11
                  * @ 6,5 SAY "McKesson"
                  REPORT FORM RDMMC TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=12
                  * @ 6,5 SAY "Cardinal Health"
                  REPORT FORM RDMCL TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=13
                  * @ 6,5 SAY "F. Dohmen/Smith Drug"
                  REPORT FORM RDMFD TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=14
                  * @ 6,5 SAY "Harvard Drug"
                  REPORT FORM RDMH TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE MCUSTNO=15
                  * @ 6,5 SAY "Bellco Drug/M&D"
                  REPORT FORM RDMBD TO PRINTER NOCONSOLE FOR WITHINRANGE()
                CASE LASTKEY()=27.OR.MCUSTNO=16
              ENDCASE

              CLEARSTATUS()
// end

//          GETMFGRANGE()
//          IF LASTKEY() != 27 .OR. MCUSTNO=7
//            MYORN=.F.
//            @ 23,0 SAY 'Is the printer ready? [y/n]' GET MYORN PICT 'Y'
//            READ
//            IF MYORN
//              @ 23,0 CLEAR
//              @ 23,0 SAY 'Printing.  Please wait...'
//              SELECT RDINV
//              SET ORDER TO TAG RDINV  &&,RDBITEM
//              GO TOP

//              @ 6,5 SAY "Master Inventory Listing..."
//              REPORT FORM RDMAST TO PRINTER NOCONSOLE FOR WITHINRANGE()
//              @ 23,0 CLEAR
//              @ 23,0 SAY "Report Printing Completed.  Press Any Key..."
//              INKEY(0)
//              @ 6,5 SAY "                           "
//            ENDIF
//          ENDIF

            MSGBOX("Report Printing Completed.  Press Any Key...")
		endif
		endif
      CASE MREPTNO=3
        DO WHILE .T.
          *CLOS ALL
          DO BORDER WITH .T.
          @ 6,5 TO 18,45 DOUBLE
          SET CONFIRM ON
          @ 5,5 SAY 'Enter Order # to Print...'
          MPO=SPACE(8)
          @ 10,6 SAY "Purchase Order:" GET MPO PICTURE "!!!!!!!!"
          READ
          @ 8,6 CLEAR TO 17,44
          IF LASTKEY()=27
            EXIT
          ENDIF
          MFNAME=DISKDIR+"\PO\"+TRIM(MPO)+'.BPO'
          IF FILE(MFNAME)
            USE &MFNAME ALIAS TEMP NEW
          ELSE
            @ 23,0 CLEAR
            @ 23,0 SAY 'Purchase order not on file.  Press any key...'
            INKEY (0)
            @ 23,0 CLEAR
            LOOP
          ENDIF
          @ 23,0 CLEAR
          @ 23,0 SAY 'Loading purchase order.  Please wait...'
          INDEX ON RITEM TO TEMP1
          SET INDEX TO TEMP1
          MYORN=.T.
          @ 23,0 CLEAR
          @ 23,0 SAY 'Is the printer ready? [y/n]' GET MYORN PICT "y"
          READ
          @ 23,0 CLEAR
          IF MYORN
            @ 23,0 SAY 'Printing.  Please wait...'
            M12PITCH = CHR(27)+'M'
            MNORMAL  = CHR(27)+'P'
            SET PRINT ON
            SET CONSOLE OFF
            ?? M12PITCH
            SET CONSOLE ON
            ?'PURCHASE ORDER: '+MPO
            SET PRINT OFF
            REPORT FORM RDBORD FOR .NOT.DELETED() TO PRINTER NOCONSOLE
            SET PRINT ON
            SET CONSOLE OFF
            ? MNORMAL
            SET PRINT OFF
            SET CONSOLE ON
            @ 23,0 CLEAR
            @ 23,0 SAY "Report Printing Completed.  Press Any Key..."
            INKEY(0)
          ENDIF
          USE
        ENDDO

      CASE MREPTNO=4
        DO WHILE .T.
          DO BORDER WITH .T.
          @ 6,5 TO 18,45 DOUBLE
          SET CONFIRM ON
          @ 5,5 SAY 'Enter Order # to Print...'
          MPO=SPACE(8)
          @ 10,6 SAY "Purchase Order:" GET MPO PICTURE "!!!!!!!!"
          READ
          @ 8,6 CLEAR TO 17,44
          IF LASTKEY()=27
            EXIT
          ENDIF
          MFNAME=DISKDIR+"\PO\"+TRIM(MPO)+'.MPO'
          IF FILE(MFNAME)
            USE &MFNAME ALIAS TEMP NEW
          ELSE
            @ 23,0 CLEAR
            @ 23,0 SAY 'Purchase order not on file.  Press any key...'
            INKEY (0)
            @ 23,0 CLEAR
            LOOP
          ENDIF
          @ 23,0 CLEAR
          @ 23,0 SAY 'Loading purchase order.  Please wait...'
          INDEX ON RITEM TO TEMP1
          SET INDEX TO TEMP1
          MYORN=.T.
          @ 23,0 CLEAR
          @ 23,0 SAY 'Is the printer ready? [y/n]' GET MYORN PICT "y"
          READ
          @ 22,0 CLEAR
          IF MYORN
            @ 23,0 SAY 'Printing.  Please wait...'
            M12PITCH = CHR(27)+'M'
            MNORMAL  = CHR(27)+'P'
            SET PRINT ON
            SET CONSOLE OFF
            ?? M12PITCH
            SET CONSOLE ON
            ?'PURCHASE ORDER: '+MPO
            SET PRINT OFF
            REPORT FORM RDMCORD FOR .NOT.DELETED() TO PRINTER NOCONSOLE
            SET PRINT ON
            SET CONSOLE OFF
            ? MNORMAL
            SET PRINT OFF
            SET CONSOLE ON
            @ 23,0 CLEAR
            @ 23,0 SAY "Report Printing Completed.  Press Any Key..."
            INKEY(0)
          ENDIF
          USE
        ENDDO
      ENDCASE
      @ 23,0 CLEAR
      @ 6,5 TO 23,45 DOUBLE
    ELSE
      EXIT
    ENDIF
  ENDDO
  SELECT RDINV
  GO TOP
RETURN
//====================== RD Transmission =====================
PROCEDURE RDTRANS
  LOCAL MCUST[17]
  * MCUST[1]='Bindley Western'
  * MCUST[2]='Kinray'
  * MCUST[3]='H.D. Smith'
  * MCUST[4]='Burlington/Amerisource Drug'
  * MCUST[5]='Walker Drug (Legacy)'
  * MCUST[6]='Schein'
  * MCUST[7]='Bergen-Brunswig'
  * MCUST[8]='Rochester/Dakota'
  * MCUST[9]='Walsh'
  * MCUST[10]='D & K Drug'
  * MCUST[11]='Walgreens'
  * MCUST[12]='McKesson'
  * MCUST[13]='Cardinal Health'
  * MCUST[14]='F. Dohmen/Smith Drug'
  * MCUST[15]='Harvard Drug'
  * MCUST[16]='Bellco Drug/M&D'
			MCUST[1]='Bindley Western'
			MCUST[2]='Walsh'
			MCUST[3]='Bellco Drug/M&D'
			MCUST[4]='Bergen-Brunswig'
			MCUST[5]='Burlington/Amerisource/NC Mutual/Longs'
			MCUST[6]='Cardinal Health'
			MCUST[7]='D & K Drug'
			MCUST[8]='F. Dohmen/Smith Drug'
			MCUST[9]='H.D. Smith'
			MCUST[10]='Harvard Drug'
			MCUST[11]='Kinray'
			MCUST[12]='McKesson'
			MCUST[13]='Schein'
			MCUST[14]='Rochester/Dakota'
			MCUST[15]='Walgreens'
			MCUST[16]='Walker Drug (Legacy)'
  MCUST[17]='Return to Main Menu'
  MCUSTNO=1

  @ 5,47 SAY '...Select Customer to Prepare'
  @ 5,5 SAY 'ษอออออออออออออออออออออออออออออออออออออออป'
  @ 6,5 SAY 'บ                                       บ'

  @ 6,6 CLEAR TO 22,44
  @ 15,47 SAY "ENTER = SELECT"
  @ 16,47 SAY "ESC   = QUIT  "
  MCUSTNO=ACHOICE(6,6,22,44,MCUST,.T.,'',MCUSTNO)

			DO CASE
				CASE MCUSTNO = 1
					MCUSTNO = 1
				CASE MCUSTNO = 2
					MCUSTNO = 9
				CASE MCUSTNO = 3
					MCUSTNO = 16
				CASE MCUSTNO = 4
					MCUSTNO = 7
				CASE MCUSTNO = 5
					MCUSTNO = 4
				CASE MCUSTNO = 6
					MCUSTNO = 13
				CASE MCUSTNO = 7
					MCUSTNO = 10
				CASE MCUSTNO = 8
					MCUSTNO = 14
				CASE MCUSTNO = 9
					MCUSTNO = 3
				CASE MCUSTNO = 10
					MCUSTNO = 15
				CASE MCUSTNO = 11
					MCUSTNO = 2
				CASE MCUSTNO = 12
					MCUSTNO = 12
				CASE MCUSTNO = 13
					MCUSTNO = 6
				CASE MCUSTNO = 14
					MCUSTNO = 8
				CASE MCUSTNO = 15
					MCUSTNO = 11
				CASE MCUSTNO = 16
					MCUSTNO = 5
			ENDCASE

  IF LASTKEY()=27.OR.MCUSTNO=17
    RETURN
  ENDIF
  MENUTO := 0
  booCancel := .F.

  IF MCUSTNO=16
    SSCREEN=SAVESCREEN(6,14,17,65)
    @ 6,14 CLEAR TO 17,65
    @ 7,15,16,64 BOX B_SINGLE+" "
    @ 11,22 SAY "Please select a company to export:"
    @ 13,25 PROMPT "Bellco Drug     "
    @ 13,44 PROMPT "Morris & Dickson"
    MENU TO MENUTO
    RESTSCREEN(6,14,17,65,SSCREEN)
	IF MENUTO == 0
		booCancel := .T.
	ENDIF
  ENDIF
  
  IF MCUSTNO=4
    SSCREEN=SAVESCREEN(6,14,17,65)
    @ 6,14 CLEAR TO 17,65
    @ 7,15,16,64 BOX B_SINGLE+" "
    @ 11,22 SAY "Please select a company to export:"
	@ 13,18 PROMPT "Amerisource"
	@ 13,31 PROMPT "Burlington"
	@ 13,43 PROMPT "NC Mutual"
	@ 13,54 PROMPT "Longs"
    MENU TO MENUTO
    RESTSCREEN(6,14,17,65,SSCREEN)
	IF MENUTO == 0
		booCancel := .T.
	ENDIF
  ENDIF

  IF MCUSTNO==8
    SSCREEN=SAVESCREEN(6,14,17,65)
    @ 6,14 CLEAR TO 17,65
    @ 7,15,16,64 BOX B_SINGLE+" "
    @ 11,22 SAY "Please select a company to export:"
    @ 13,28 PROMPT "Rochester"
    @ 13,41 PROMPT "Dakota   "
    MENU TO MENUTO
    RESTSCREEN(6,14,17,65,SSCREEN)
	IF MENUTO == 0
		booCancel := .T.
	ENDIF
  ENDIF

  IF MCUSTNO == 13 .OR. MCUSTNO == 1 .OR. MCUSTNO == 7 .OR. MCUSTNO == 12
	sScreen := SAVESCREEN()
	@ 7,15,16,64 BOX B_SINGLE+" "
	@ 11,17 SAY "Type of Export:"
	*@ 13,33 PROMPT "Old"
	@ 13,33 PROMPT "AD Only"
	@ 13,43 PROMPT "All"
	MENU TO MENUTO
	RESTSCREEN(,,,,sScreen)
	IF MENUTO == 0
		booCancel := .T.
	ENDIF
  ENDIF
  IF MCUSTNO == 14
	sScreen := SAVESCREEN()
	@ 7,15,16,64 BOX B_SINGLE+" "
	@ 11,17 SAY "Please select a company to export:"
	@ 13,30 PROMPT "F. Dohmen"
	@ 13,41 PROMPT "Smith Drug"
	MENU TO MENUTO
	RESTSCREEN(,,,,sScreen)
	IF MENUTO == 0
		booCancel := .T.
	ENDIF
  ENDIF
    
  MYORN=.F.
  IF !booCancel
	@ 23,0 SAY ' Continue? [y/n] '  COLOR "W+/B"
	@ 23,17 GET MYORN PICT 'Y' COLOR "W+/BG"
	READ
  ENDIF
  
  IF MYORN
    @ 23,0 CLEAR
    * @ 5,5  SAY '                             '
    @ 23,0 SAY ' Processing.  Please wait... ' COLOR "W+/R"
    SELECT RDINV
    SET ORDER TO TAG RDINV  &&,RDBITEM
    GO TOP
	IF MENUTO == 1
		strAD = "(AD Only)"
	ELSE
		strAD = "(ALL)"
	ENDIF
    DO CASE
      CASE MCUSTNO=1
		@ 10,44 SAY "Bindley Western "+strAD COLOR "G+*/N"
		strFileName = "BWTRAN.TXT"
		DO BWTRANS WITH MENUTO
		SUCCESS := COPYFILE(strFileName)
      CASE MCUSTNO=2
        @ 10,44 SAY "Kinray" COLOR "G+*/N"
        DO KTRANS
        // SIMON
        SUCCESS := COPYFILE("KINRAY.TXT")
      CASE MCUSTNO=3
        @ 10,44 SAY "H.D. Smith" COLOR "G+*/N"
        DO TTRANS
        // SIMON
        SUCCESS := COPYFILE("AIVREB")
      CASE MCUSTNO=4
        IF MENUTO = 1
	        @ 10,44 SAY "Amerisource Drug" COLOR "G+*/N"
			strFileName = "AMTRAN.TXT"
	        DO AMTRANS
		ELSEIF MENUTO = 2
	        @ 10,44 SAY "Burlington Drug" COLOR "G+*/N"
			strFileName = "BUTRAN.TXT"
	        DO WTRANS WITH "B"
		ELSEIF MENUTO = 3
	        @ 10,44 SAY "NC Mutual" COLOR "G+*/N"
			strFileName = "REBEL.TXT"
	        DO WTRANS WITH "N"
		ELSEIF MENUTO = 4
	        @ 10,46 SAY "Longs" COLOR "G+*/N"
			strFileName = "xfiasitrb.TXT"
	        DO WTRANS WITH "L"
		ENDIF
		SUCCESS := COPYFILE(strFileName)
      CASE MCUSTNO=5
        @ 10,44 SAY "Walker Drug (Legacy)" COLOR "G+*/N"
        DO W2TRANS
        // SIMON
        SUCCESS := COPYFILE("W2TRAN.TXT")
      CASE MCUSTNO=6
        @ 10,44 SAY "Schein" COLOR "G+*/N"
        DO NTRANS
        SUCCESS := COPY_FILE()
      CASE MCUSTNO=7
        @ 10,44 SAY "Bergen-Brunswig "+strAD COLOR "G+*/N"
        DO BTRANS WITH MENUTO
      CASE MCUSTNO=8
        IF MENUTO == 1
			@ 10,44 SAY "Rochester" COLOR "G+*/N"
			strFileName := "ROTRAN.TXT"
		ELSE
			@ 10,44 SAY "Dakota" COLOR "G+*/N"
			strFileName := "DAKOTA.TXT"
		ENDIF
		DO BNTRANS WITH strFileName
        SUCCESS := COPYFILE(strFileName)
      CASE MCUSTNO=9
        @ 10,44 SAY "Walsh" COLOR "G+*/N"
        DO BRTRANS
        // Steve
        SUCCESS := COPYFILE("rebwal.txt")
      CASE MCUSTNO=10
        @ 10,44 SAY "D & K Drug" COLOR "G+*/N"
        DO DKTRANS
        // Simon
        SUCCESS := COPYFILE("DKTRAN.TXT")
      CASE MCUSTNO=11
        @ 10,44 SAY "Walgreens" COLOR "G+*/N"
        DO WGTRANS
        // Simon
        SUCCESS := COPYFILE("C:\COMMERCE\IMPORTS\WGTRAN.TDF")
      CASE MCUSTNO=12
        @ 10,46 SAY "McKesson "+strAD COLOR "G+*/N"
        DO MCTRANS WITH MENUTO
        // Simon
        SUCCESS := COPYFILE("MCTRAN.TXT")
      CASE MCUSTNO=13
        strFileName = "CLTRAN.TXT"
        IF !booCancel
			@ 10,44 SAY "Cardinal Health "+strAD COLOR "G+*/N"
			IF MENUTO != 1
				strFileName := "CLTRANAL.TXT"
			ENDIF
			DO CLTRANS WITH MENUTO
			// Simon
			SUCCESS := COPYFILE(strFileName)
		ENDIF
      CASE MCUSTNO=14
        if !booCancel
			if menuto == 1
				@ 10,44 SAY "F. Dohmen" COLOR "G+*/N"
				sfile := "DREBELA"
			else
				@ 10,44 SAY "Smith Drug" COLOR "G+*/N"
				sfile := "REBELSD.TXT"
			endif
			
			DO FDTRANS with menuto
			// Simon
			SUCCESS := COPYFILE(sfile)
		endif
      CASE MCUSTNO=15
        @ 10,44 SAY "Harvard Drug" COLOR "G+*/N"
        DO HTRANS
        // Simon
        * SUCCESS := COPYFILE("HTRAN.TXT")
      CASE MCUSTNO=16
        IF MENUTO=1
	        @ 10,44 SAY "Bellco Drug" COLOR "G+*/N"
			strFileName = "BDTRAN.TXT"
		ELSE
			@ 10,44 SAY "Morris & Dickson" COLOR "G+*/N"
			strFileName = "REB.TXT"
		ENDIF
        DO BDTRANS WITH MENUTO
        SUCCESS := COPYFILE(strFileName)
      CASE LASTKEY()=27.OR.MCUSTNO=17
        @ 10,44 SAY '                                 '
        RETURN
    ENDCASE
    @ 23,0 CLEAR
    @ 10,44 SAY '                                 '
    IF SUCCESS
      @ 23,0 SAY " Transmission File and Mailing Disk Prepared.  Press Any Key... " COLOR "W+/GR"
    ELSE
      IF MCUSTNO == 4
        @ 23,0 SAY " Mailing Diskette Aborted,  Only Transmission File Prepared.  Press Any Key... " COLOR "W+/BG"
      ELSE
        if !booCancel
			@ 23,0 SAY " Transmission File Prepared.  Press Any Key... " COLOR "W+/G"
	    ENDIF
      ENDIF
    ENDIF
    SUCCESS := .F.
    INKEY(0)
  ENDIF
RETURN


PROCEDURE BTRANS (ADMODE)
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\BTRANS.TXT
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF .NOT.BTRANSMIT
      SKIP
      LOOP
    ENDIF
    
	  IF ADMODE != 2
		IF !CLAD
			SKIP
			LOOP
		ENDIF
	  ENDIF
    
    C01=BITEMNO                   && BERGEN ITEM NUMBER
    C02=SPACE(2)                  && RESERVED
    C03='78088'                   && SUPPLIER NUMBER
    C04=PADL(BQTYAVAIL,5,'0')     && QUANTITY AVAILABLE
    MPRICE=ROUND(BPRICE*100,0)
    C05=PADL(MPRICE,7,'0')        && PRICE
    C06='1'                       && PRICE CODE
    C07=PADL(BCASEPACK,3,'0')     && CASE PACK QUANTITY
    C08=PADL(BMINORDER,3,'0')     && MIN. ORDER QUANTITY
    C09='1'                       && MIN ORDER CODE
    C10=SPACE(11)                 && NDC CODE
    C11=SPACE(6)                  && USER DEFINED
    C12=SPACE(30)                 && TRANSMITTED DATA ONLY
    ? C01+C02+C03+C04+C05+C06+C07+C08+C09+C10+C11+C12
    SKIP
  ENDDO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
  *CLOS ALL
RETURN

PROCEDURE KTRANS
  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\KINRAY.TXT
  SET ALTERNATE ON

  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  DO WHILE .NOT.EOF()
    IF (!FTRANSMIT) .OR. (!FVALID)
      SKIP
      LOOP
    ENDIF

    *DETAIL
    IF !EMPTY(FITEMNO).AND.!EMPTY(FPRICE).AND.!EMPTY(FQTYAVAIL)
      SPRICE := STR( FPRICE )
      C01 := "REBEL     "
      C02 := PADR(ALLTRIM(FITEMNO),18," ")
      C03 := "I"
      C04 := PADL( ALLTRIM( STR( FQTYAVAIL ) ),7,"0" )
      C05 := PADL(ALLTRIM( SUBSTR(SPRICE,1,AT( ".",SPRICE )-1) ) + PADR(ALLTRIM( SUBSTR(SPRICE,AT( ".",SPRICE )+1) ),4,"0"),13,"0" )
      C06 := "000010000"
      C07 := "0"
      ?? C01+C02+C03+C04+C05+C06+C07
      ?
    ENDIF
    SKIP
  ENDDO

  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN

PROCEDURE TTRANS
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\AIVREB.
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF .NOT.TTRANSMIT
      SKIP
      LOOP
    ENDIF
    SPRICE := ALLTRIM( STR( TPRICE ) )
    MPRICE= ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )

    C01=PADL(TRIM(TITEMNO),11,"0")
    C02=PADL(TQTYAVAIL,5,'0')
    C03=PADL(MPRICE,6,'0')+"0"        && PRICE
    C04=SPACE(7)
    C05="001"
    ?? C01+C02+C03+C04+C05
    ?

    *TITEMNO+SPACE(3)                   && ITEM NUMBER
    *PADL(TCASEPACK,5,'0')     && CASE PACK QUANTITY
    *C05=SPACE(1)
    *C06=PADL(TQTYAVAIL,5,'0')     && QUANTITY AVAILABLE
    *MISSINGSPACE = SPACE(1)
    *C08=SPACE(14)
    *C09=SPACE(1)

    *C01=TITEMNO
    *C02=SUBSTR( RITEM,1,29 )
    *C03=PADL( SUBSTR( STRPRICE,1,MPOS := AT( ".",STRPRICE )-1 )+SUBSTR( STRPRICE,MPOS+2,LEN( STRPRICE ) ),6," " )+SPACE( 31 )
    *? C01+C02+C03
    SKIP
  ENDDO
  SET ALTERNATE OFF
  SET ALTERNATE TO
  SET(_SET_CONSOLE,.T.)
  *CLOS ALL
RETURN


*PROCEDURE DKTRANS
*  SELECT RDINV
*  SET ORDER TO TAG RDINV  &&,RDBITEM
*  GO TOP
*
*  SET(_SET_CONSOLE,.F.)
*  SET ALTERNATE TO EXPORT\DKTRAN.TXT
*  SET ALTERNATE ON
*  DO WHILE .NOT.EOF()
*    IF .NOT.TTRANSMIT
*      SKIP
*      LOOP
*    ENDIF
*
*    SPRICE := ALLTRIM( STR( TPRICE ) )
*    MPRICE= ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )+"00"
*
*    C01="REBEL DIST"
*    C02=PADR(PADL(TRIM(TITEMNO),11,"0"),18)
*    C03="U"
*    C04=PADL(TQTYAVAIL,7,"0")     && QUANTITY AVAILABLE
*    C05=PADL(MPRICE,9,"0")        && PRICE
*    C06="000000000"
*    C07="0"
*    ? C01+C02+C03+C04+C05+C06+C07
*
*    *TITEMNO+SPACE(3)                   && ITEM NUMBER
*    *PADL(TCASEPACK,5,'0')     && CASE PACK QUANTITY
*    *C05=SPACE(1)
*    *C06=PADL(TQTYAVAIL,5,'0')     && QUANTITY AVAILABLE
*    *MISSINGSPACE = SPACE(1)
*    *C08=SPACE(14)
*    *C09=SPACE(1)
*
*    *C01=TITEMNO
*    *C02=SUBSTR( RITEM,1,29 )
*    *C03=PADL( SUBSTR( STRPRICE,1,MPOS := AT( ".",STRPRICE )-1 )+SUBSTR( STRPRICE,MPOS+2,LEN( STRPRICE ) ),6," " )+SPACE( 31 )
*    *? C01+C02+C03
*    SKIP
*  ENDDO
*  SET ALTERNATE OFF
*  SET ALTERNATE TO
*  SET(_SET_CONSOLE,.T.)
*  *CLOS ALL
*RETURN


PROCEDURE MCTRANS (ADMODE)
  LOCAL RCOUNT := 0
  mREF := "S"

  @23,0 CLEAR
  @23,0 SAY "(S)tandard, (R)efrigerated, or (I)njectable? "
  @23,45 GET mREF PICTURE "@!" VALID mREF $ "SIR"
  READ

  if LASTKEY() != 27
  
	SELECT RDINV
	SET ORDER TO TAG RDINV  &&,RDBITEM
	GO TOP

	SET(_SET_CONSOLE,.F.)
	SET ALTERNATE TO EXPORT\MCTRAN.TXT
	SET ALTERNATE ON

	* HEADER
	IF mREF="I"
		?? "H65087Rebel Distributors            "+SUBSTR(DTOS(DATE()),3)+SPACE(37)
	ELSE
		IF mREF="R"
			?? "H65167Rebel Distributors            "+SUBSTR(DTOS(DATE()),3)+SPACE(37)
		ELSE
			?? "H65166Rebel Distributors            "+SUBSTR(DTOS(DATE()),3)+SPACE(37)
		ENDIF
	ENDIF

	DO WHILE .NOT.EOF()
		IF .NOT.MCTRANSMIT
		SKIP
		LOOP
		ENDIF

		IF MCREFTYPE != mREF
		SKIP
		LOOP
		ENDIF

	  IF ADMODE != 2
		IF !CLAD
			SKIP
			LOOP
		ENDIF
	  ENDIF

		*IF .NOT.MCREF .AND. mREF="R"
		*  SKIP
		*  LOOP
		*ENDIF
	           
		*DETAIL
		IF !EOF().AND.!EMPTY(MCITEMNO).AND.!EMPTY(MCPRICE).AND.!EMPTY(MCQTYAVAIL)
		RCOUNT++
		SPRICE := STR( MCPRICE )
		C01 := "D"
		C02 := PADL(ALLTRIM(MCITEMNO),7)
		C03 := "     "
		C04 := MCORDUNIT  && EA, CT, OR CS
		C05 := PADL(ALLTRIM( SUBSTR(SPRICE,1,AT( ".",SPRICE )-1) ),5,"0" )
		C06 := PADR(ALLTRIM( SUBSTR(SPRICE,AT( ".",SPRICE )+1) ),4,"0" )
		C07 := PADL( ALLTRIM( STR( MCQTYAVAIL ) ),5,"0" )
		C08 := SPACE(50)
		? C01+C02+C03+C04+C05+C06+C07+C08
		ENDIF
		SKIP
	ENDDO

	*FOOTER
	? "T"+IIF(RCOUNT>0,PADL(ALLTRIM(STR(RCOUNT)),7),SPACE(7))+PADL(DTOC(DATE())+" - Rebel Distributors Corporation (818) 865-6880",72)

	SET ALTERNATE TO
	SET ALTERNATE OFF
	SET(_SET_CONSOLE,.T.)
  END IF
RETURN


PROCEDURE AMTRANS
  LOCAL RCOUNT := 0
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\AMTRAN.TXT
  SET ALTERNATE ON

  DO WHILE .NOT.EOF()
    IF .NOT.WTRANSMIT
      SKIP
      LOOP
    ENDIF

    *DETAIL
    IF !EMPTY(WITEMNO).AND.!EMPTY(WPRICE).AND.!EMPTY(WQTYAVAIL)
      SPRICE := STR( WPRICE*12 )
      C01 := "REBEL     "
      C02 := PADR(PADL(ALLTRIM(WITEMNO),11,"0"),15)
      C03 := "N"
      C04 := PADL(ALLTRIM( SUBSTR(SPRICE,1,AT( ".",SPRICE )-1) ) + ALLTRIM( SUBSTR(SPRICE,AT( ".",SPRICE )+1) ),8,"0" )+"0"
      C05 := PADL( ALLTRIM( STR( WQTYAVAIL ) ),7,"0" )
      ?? C01+C02+C03+C04+C05
      ?
    ENDIF
    SKIP
  ENDDO

  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN


PROCEDURE WTRANS( MODE )
  LOCAL RCOUNT := 0
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  
  DO CASE
	CASE MODE = "B"
		SET ALTERNATE TO EXPORT\BUTRAN.TXT
	CASE MODE = "N"
		SET ALTERNATE TO EXPORT\REBEL.TXT
	CASE MODE = "L"
		SET ALTERNATE TO EXPORT\XFIASITRB.TXT
  END CASE
  
  SET ALTERNATE ON

  DO WHILE .NOT.EOF()
    IF .NOT.WTRANSMIT
      SKIP
      LOOP
    ENDIF

    *DETAIL
    IF !EMPTY(WITEMNO).AND.!EMPTY(WPRICE).AND.!EMPTY(WQTYAVAIL)
      * RCOUNT++
      * SPRICE := STR( WPRICE )
      * C01 := "RD"
      * C02 := "ST"
      * C03 := "1"
      * C04 := PADL(ALLTRIM(WITEMNO),11,"0")
      * C05 := "     "
      * C06 := "1"
      * C07 := PADL( ALLTRIM( STR( WQTYAVAIL ) ),4,"0" )
      * C08 := PADL(ALLTRIM( SUBSTR(SPRICE,1,AT( ".",SPRICE )-1) ) + ALLTRIM( SUBSTR(SPRICE,AT( ".",SPRICE )+1) ),6,"0" )
      * C09 := "0000"
      * C10 := "0001"
      * C11 := MEMOLINE(RITEM,26,1)
      * ? C01+C02+C03+C04+C05+C06+C07+C08+C09+C10+C11

      SPRICE := STR( WPRICE*12 )
      C01 := "REBEL     "
      C02 := PADR(PADL(ALLTRIM(WITEMNO),11,"0"),15)
      
      IF MODE="N"
		C03 := "U"
	  ELSE
		C03 := "N"
      ENDIF
      
      C04 := PADL(ALLTRIM( SUBSTR(SPRICE,1,AT( ".",SPRICE )-1) ) + ALLTRIM( SUBSTR(SPRICE,AT( ".",SPRICE )+1) ),8,"0" )+"0"
      C05 := PADL( ALLTRIM( STR( WQTYAVAIL ) ),7,"0" )
      ?? C01+C02+C03+C04+C05
      ?
    ENDIF
    SKIP
  ENDDO

  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)

  if mode = "L"
  	*RUN ("DIR")
  	X = SAVESCREEN()
*        RUN ("xcopy export\xfiasitr.txt export\xfiiasitrb.txt")
  	RESTSCREEN(,,,,X) 
  endif

RETURN

PROCEDURE W2TRANS
  LOCAL RCOUNT := 0
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\WTRAN.TXT
  SET ALTERNATE ON

  DO WHILE .NOT.EOF()
    IF .NOT.WTRANSMIT
      SKIP
      LOOP
    ENDIF

    *DETAIL
    IF !EMPTY(WITEMNO).AND.!EMPTY(WPRICE).AND.!EMPTY(WQTYAVAIL)
      RCOUNT++
      SPRICE := STR( WPRICE )
      C01 := "RD"
      C02 := "ST"
      C03 := "1"
      C04 := PADL(ALLTRIM(WITEMNO),11,"0")
      C05 := "     "
      C06 := "1"
      C07 := PADL( ALLTRIM( STR( WQTYAVAIL ) ),4,"0" )
      C08 := PADL(ALLTRIM( SUBSTR(SPRICE,1,AT( ".",SPRICE )-1) ) + ALLTRIM( SUBSTR(SPRICE,AT( ".",SPRICE )+1) ),6,"0" )
      C09 := "0000"
      C10 := "0001"
      C11 := MEMOLINE(RITEM,26,1)
      ?? C01+C02+C03+C04+C05+C06+C07+C08+C09+C10+C11

      * C01 := "76750     "
      * C02 := PADR(ALLTRIM(RNDC),15)
      * C03 := "N"
      * C04 := PADL(TRANSFORM(WPRICE,"99999.999"),9,"0")
      * C05 := PADL( ALLTRIM( STR( WQTYAVAIL ) ),7,"0" )
      * ?? C01+C02+C03+C04+C05
      ?
    ENDIF
    SKIP
  ENDDO

  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN



PROCEDURE NTRANS
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\REBEL
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF .NOT. NTRANSMIT
      SKIP
      LOOP
    ENDIF


    *C01 = PADR( PADL(ALLTRIM(           NITEMNO),11,"0"),18," " )
    *C02 = PADL( ALLTRIM( STR( NQTYAVAIL ) ), 7,"0" )
    *C03 = PADL( ALLTRIM( STR(    NPRICE ) ), 8,"0" )
    *C03 = SUBSTR(C03,1,5) + SUBSTR(C03,7,2)+"00"
    *?? "REBEL     "+C01+"U"+C02+C03+"0000000000"
    *?  ""
    IF AT("-",NITEMNO) > 0
      SNDC   := SUBSTR(NITEMNO,1,AT("-",NITEMNO)-1)+SUBSTR(NITEMNO,AT("-",NITEMNO)+1)
    ELSE
      SNDC   := ALLTRIM(NITEMNO)
    ENDIF

    SPRICE := STR(NPRICE)
    MPRICE := ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )
    C01 := "REBEL     "
    C02 := PADR(PADL(ALLTRIM(SNDC),11,"0"),18)
    C03 := "U"
    C04 := PADL( ALLTRIM( STR( NQTYAVAIL ) ), 7,"0" )
    C05 := PADL(MPRICE,11,"0" )+"00"
    C06 := "000000000"
    C07 := "1"
    ?? C01+C02+C03+C04+C05+C06+C07
    ?
    SKIP
  ENDDO
  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN


PROCEDURE BWTRANS (ADMODE)
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\BWTRAN.TXT
  
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF .NOT. BWTRANSMIT
      SKIP
      LOOP
    ENDIF

  	  IF ADMODE != 2
		IF !CLAD
			SKIP
			LOOP
		ENDIF
	  ENDIF

    
    SPRICE := STR( BWPRICE )
    MPRICE := ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )

		C01 := "7190"
		C02 := PADL(AllTrim(BWITEMNO),11,"0")
		C03 := " "
		C04 := PADL(BWIN,6,"0")
		C05 := " "
		C06 := SUBSTR(RITEM,1,30)
		C07 := PADL(BWQTYAVAIL,5," ")
		C08 := " "
		C09 := PADL(MPRICE,5," ")

		*C01 = PADR(PADL(ALLTRIM(BWITEMNO),15,"0"),22," ")
		*C02 = " "+PADR(RITEM,30)
		*C03 = PADL( ALLTRIM( STR( BWQTYAVAIL ) ), 5,"0" )
		*C04 = " "
		*C05 = PADL( ALLTRIM(SPRICE), 6,"0" )
		*C06 = SPACE(14)
		?? C01+C02+C03+C04+C05+C06+C07+C08+C09
    ?  ""
    SKIP
  ENDDO
  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN


PROCEDURE BNTRANS( strFileName )
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  pstrFileName := "EXPORT\"+strFileName
  SET ALTERNATE TO (pstrFileName)
  *EXPORT\ROTRAN.TXT
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF .NOT. JTRANSMIT
      SKIP
      LOOP
    ENDIF

    SPRICE := STR( JPRICE )
    MPRICE= ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )

    IF strFileName == "ROTRAN.TXT"
		C01 = PADL(ALLTRIM(JITEMNO),11,"0")
		C02 = PADL(SUBSTR(RITEM,1,30),30)
		C03 = "0000000001"
		C04 = PADL( ALLTRIM( STR( JQTYAVAIL ) ), 6,"0" )
		C05 = PADL( MPRICE,7,"0" )+"0"
	ELSE
		C01 = PADL(ALLTRIM(JITEMNO),11,"0")
		C02 = PADL( MPRICE,7,"0" )+"0"
		C03 = PADL( ALLTRIM( STR( JQTYAVAIL ) ), 6,"0" )
		C04 = PADR(SUBSTR(RITEM,1,39),39)
		C05 = ""
	ENDIF
    ?? C01+C02+C03+C04+C05
    ?  ""
    SKIP
  ENDDO
  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN



PROCEDURE brTRANS
// Steve 9/6/95 - Added

  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\rebwal.txt
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF .NOT.brTRANSMIT
      SKIP
      LOOP
    ENDIF
    SPRICE := ALLTRIM( STR( brPRICE ) )
    MPRICE= ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )
    strItem := padl(TRIM(brITEMNO),13,"0")
    *C01=SUBSTR(strItem,2)
    *C02=SPACE(43)
    *C03 = PADL( ALLTRIM( STR( brQTYAVAIL ) ), 5,"0" )
    *C04=SPACE(1)
    *C05=PADL(MPRICE,6,"0")        && PRICE
    
    C01 := "6315"
    C02 := strItem
    C03 := PADR(RITEM,33)
    C04 := PADL(MPRICE,7,"0")
    C05 := PADL( ALLTRIM( STR( brQTYAVAIL ) ), 5,"0" )
    C06 := SPACE(18)
    ?? C01+C02+C03+C04+C05+C06
	?
    SKIP

  ENDDO
  SET ALTERNATE OFF
  SET ALTERNATE TO
  SET(_SET_CONSOLE,.T.)
RETURN


PROCEDURE DKTRANS
  // Simon 9/18/95 - Added
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO EXPORT\DKTRAN.TXT
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF .NOT.BETRANSMIT
      SKIP
      LOOP
    ENDIF
    SPRICE    := ALLTRIM( STR( BEPRICE ) )
    MPRICE    := ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )
    *MEXTPRICE := ALLTRIM( STR( BEPRICE*BEQTYAVAIL ) )
    *MEXTPRICE := ALLTRIM( SUBSTR( MEXTPRICE,1,AT( ".",MEXTPRICE )-1 )+SUBSTR( MEXTPRICE,AT( ".",MEXTPRICE )+1,LEN( MEXTPRICE ) ) )

    C01=PADL(TRIM(BEITEMNO),11,"0")
    C02=PADR(RITEM,30)
    C03=PADL(ALLTRIM(STR(BEQTYAVAIL)),7,"0")
    C04=PADL(MPRICE,9,"0")        && PRICE
    *C05=PADL(MEXTPRICE,9,"0")        && PRICE
    ?? C01+C02+C03+C04
    ?
    SKIP

  ENDDO
  SET ALTERNATE OFF
  SET ALTERNATE TO
  SET(_SET_CONSOLE,.T.)
RETURN


PROCEDURE WGTRANS
  LOCAL CNT := 0, booNew := .F., MENUTO := ""

  * sScreen := SAVESCREEN()
  * @ 7,15,16,64 BOX B_SINGLE+" "
  * @ 11,17 SAY "Use Old or New Export?"
  * @ 13,35 PROMPT "Old"
  * @ 13,40 PROMPT "New"
  * MENU TO MENUTO
  * RESTSCREEN(,,,,sScreen)

  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  SET ALTERNATE TO C:\COMMERCE\IMPORTS\WGTRAN.TDF
  SET ALTERNATE ON

  * PRINT THE FOLLOWING EDI HEADERS
  ?? "864   00000000000                                       WALGREENS                                                003020                       X                                             P"
  ?  "864   BMG0020000000"
  * ?  "864   DTM00300000"
  * ?  "864   N1 00400000"
  * ?  "864   N2 00500000"
  * ?  "864   N3 00600000"
  * ?  "864   N4 00700000"
  * ?  "864   REF00800000"
  * ?  "864   PER00900000"
  ?  "864   MIT01000000X"
  * ?  "864   N1 01100000"
  * ?  "864   N2 01200000"
  * ?  "864   N3 01300000"
  * ?  "864   N4 01400000"
  * ?  "864   REF01500000"
  * ?  "864   PER01600000"

  DO WHILE .NOT.EOF()
    IF .NOT. WGTRANSMIT
      SKIP
      LOOP
    ENDIF

    MPRICE := ALLTRIM(STR(WGPRICE))

    C01 := "05"
    C02 := PADL(ALLTRIM(WGITEMNO),6,"0")
    C03 := "066623"  //Diverter Number
    C04 := PADL(ALLTRIM(STR(WGCASEPACK)),4,"0")
    C05 := "EA  "
    C06 := "0001"
    C07 := PADL(SUBSTR( MPRICE,1,AT( ".",MPRICE )-1 ),5,"0")
    C08 := PADR(SUBSTR( MPRICE,AT( ".",MPRICE )+1),4,"0")
    C09 := PADL( ALLTRIM( STR( WGQTYAVAIL ) ), 6,"0" )      
    C10 := "0001"
    C11 := StripAst(PADR(RITEM,33))
    C12 := PADR(RNDC,10)
    C13 := "ND"
    C14 := PADR(PADL(ALLTRIM(FDITEMNO),11,"0"),14," ")

    * IF MENUTO == 2

    ? "864   MSG01700000"+C01+C02+C03+C04+C05+C06+C07+C08+C09+C10+C11+C12+C13+C14

    * ELSE
    *   ? "864   MSG01700000"+C01+C02+C03+C04+C05+C06+C07+C08+C09+C10+C11+C12
    * ENDIF

    * ?? C01+C02+C03+C04+C05+C06+C07+C08+C09+C10+C11+C12+C13+C14
    * ?
    CNT += 1
    SKIP
  ENDDO
  * SUMMARY
  ? "864   MSG0170000001"+SUBSTR(DTOS(DATE()),3)+PADL(LTRIM(STR(CNT)),6,"0")+"066623"+PADR("REBEL DISTRIBUTORS",30)+SPACE(38)
  ? ""
  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN

PROCEDURE CLTRANS( MODE )
 
	SELECT RDINV
	SET ORDER TO TAG RDINV  &&,RDBITEM
	GO TOP

	SET(_SET_CONSOLE,.F.)
	*IF MODE = 1
	*	  SET ALTERNATE TO EXPORT\CLTRAN.TXT
	*ELSE
		IF MODE = 1
		  SET ALTERNATE TO EXPORT\OFFERS.TXT
		ELSE
		  SET ALTERNATE TO EXPORT\Cardinal\OFFERS.TXT
		ENDIF
	*ENDIF
	SET ALTERNATE ON
	DO WHILE .NOT.EOF()
	  IF MODE != 2
		IF !CLTRANSMIT .OR. !CLAD
			SKIP
			LOOP
		ENDIF
	  ELSE
		IF !CLTRANSMIT
			SKIP
			LOOP
		ENDIF
	  ENDIF
	  MPRICE := STR(CLPRICE)
	  
	*  IF MODE = 1
	*		C01="90115"
	*		C02=PADL(ALLTRIM(CLITEMNO),7,"0")
	*		C03=PADL(LTRIM(SUBSTR(MPRICE,1,AT(".",MPRICE)-1)+SUBSTR(MPRICE,AT(".",MPRICE)+1)),7,"0")
	*		C04=PADL(CLQTYAVAIL,5,'0')     && QUANTITY AVAILABLE
    *                    C05=""
	*	ELSE
			IF MODE = 1
				C01="90115     "
				C02=PADR(ALLTRIM(CLITEMNO),15," ")
				C03="I"
				C04=PADL(ALLTRIM(SUBSTR(MPRICE,1,AT(".",MPRICE)-1)+SUBSTR(MPRICE,AT(".",MPRICE)+1)),9,"0")
				C05=PADL(CLQTYAVAIL,7,"0")     && QUANTITY AVAILABLE
			ELSE
				C01="98115     "
				C02=PADR(ALLTRIM(CLITEMNO),15," ")
				C03="I"
				C04=PADL(ALLTRIM(SUBSTR(MPRICE,1,AT(".",MPRICE)-1)+SUBSTR(MPRICE,AT(".",MPRICE)+1)),9,"0")
				C05=PADL(CLQTYAVAIL,7,"0")     && QUANTITY AVAILABLE
			ENDIF
	*	ENDIF
      ?? C01+C02+C03+C04+C05
	  ?
	  SKIP
	ENDDO
	SET ALTERNATE OFF
	SET(_SET_CONSOLE,.T.)
	*CLOS ALL
RETURN

PROCEDURE FDTRANS( mode )
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  if mode == 1
	SET ALTERNATE TO EXPORT\DREBELA.
  else
    set alternate to export\rebelsd.txt
  endif
  
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF !FDTRANSMIT
      SKIP
      LOOP
    ENDIF
    MPRICE := STR(FDPRICE)

    IF MODE == 1
		C01="RE"
		C02=SPACE(8)
		C03=PADL(ALLTRIM(FDITEMNO),11,"0")
		C04=SPACE(7)
		C05="u"
		C06=PADL(FDQTYAVAIL,7,'0')     && QUANTITY AVAILABLE
		C07=PADL( LTRIM( SUBSTR( MPRICE,1,AT(".",MPRICE)-1 ) + PADR( SUBSTR( MPRICE,AT(".",MPRICE)+1),4,"0" ) ),9,"0" )
		C08=REPLICATE("0",10)
    else
		C01="REBEL"
		C02=SPACE(5)
		C03=PADL(ALLTRIM(FDITEMNO),11,"0")
		C04=SPACE(7)
		C05="U"
		C06=PADL(FDQTYAVAIL,7,'0')     && QUANTITY AVAILABLE
		C07=PADL( LTRIM( SUBSTR( MPRICE,1,AT(".",MPRICE)-1 ) + PADR( SUBSTR( MPRICE,AT(".",MPRICE)+1),4,"0" ) ),9,"0" )
		C08=REPLICATE("0",10)
    endif
    ?? C01+C02+C03+C04+C05+C06+C07+C08
    ?
    SKIP
  ENDDO
  SET ALTERNATE OFF
  if mode == 1
	ORGFILE := DISKDIR+"\EXPORT\DREBELA."
	FILE1 :=   DISKDIR+"\EXPORT\DREBOA."
	FILE2 := DISKDIR+"\EXPORT\DREBRA."
	COPY FILE &ORGFILE TO &FILE1
	COPY FILE &ORGFILE TO &FILE2
  endif
  SET(_SET_CONSOLE,.T.)
  *CLOS ALL
RETURN


PROCEDURE HTRANS
  Local FFILE
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)

  FFILE := "EXPORT\"+strHarvardVendorNumber+".DAT"
  SET ALTERNATE TO (FFILE)
  SET ALTERNATE ON

  DO WHILE .NOT.EOF()
    IF !HTRANSMIT
      SKIP
      LOOP
    ENDIF
    MPRICE := TRIM(STR(HPRICE*12))
    C01=PADR(strHarvardVendorNumber,10)
    C02=PADR(ALLTRIM(HITEMNO),15)
    C03="N"
    C04=PADL( LTRIM( SUBSTR( MPRICE,1,AT(".",MPRICE)-1 ) + PADR( SUBSTR( MPRICE,AT(".",MPRICE)+1),3,"0" ) ),9,"0" )
    C05=PADL(HQTYAVAIL,7,'0')     && QUANTITY AVAILABLE
    ?? C01+C02+C03+C04+C05
    ?
    SKIP
  ENDDO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
  *CLOS ALL
RETURN


PROCEDURE BDTRANS( mode )
  SELECT RDINV
  SET ORDER TO TAG RDINV  &&,RDBITEM
  GO TOP

  SET(_SET_CONSOLE,.F.)
  if mode = 1
	SET ALTERNATE TO EXPORT\BDTRAN.TXT
  else
	SET ALTERNATE TO EXPORT\REB.TXT
  endif
  
  SET ALTERNATE ON
  DO WHILE .NOT.EOF()
    IF !BDTRANSMIT .OR. !BDBWUPDATE
      SKIP
      LOOP
    ENDIF

    IF AT("-",BDITEMNO) > 0
      SNDC   := SUBSTR(BDITEMNO,1,AT("-",BDITEMNO)-1)+SUBSTR(BDITEMNO,AT("-",BDITEMNO)+1)
    ELSE
      SNDC   := ALLTRIM(BDITEMNO)
    ENDIF

    SPRICE := STR(BDPRICE)
    MPRICE := ALLTRIM( SUBSTR( SPRICE,1,AT( ".",SPRICE )-1 )+SUBSTR( SPRICE,AT( ".",SPRICE )+1,LEN( SPRICE ) ) )

	IF mode = 1
		C01 := "REBEL     "
		C02 := PADR(PADL(ALLTRIM(SNDC),11,"0"),18)
		C03 := "U"
		C04 := PADL( ALLTRIM( STR( BDQTYAVAIL ) ), 7,"0" )
		C05 := PADL(MPRICE,11,"0" )+"00"
		C06 := "000000000"
		C07 := "1"
		?? C01+C02+C03+C04+C05+C06+C07
	else
		C01 := PADL(AllTrim(BWITEMNO),11,"0")
		C02 := PADL(MPRICE,7,"0")
		C03 := PADL(BWQTYAVAIL,5,"0")
		?? C01+C02+C03
	endif
    ?
    SKIP
  ENDDO
  SET ALTERNATE TO
  SET ALTERNATE OFF
  SET(_SET_CONSOLE,.T.)
RETURN


PROCEDURE BORDER
  PARAMETER CLRSCN
  IF CLRSCN
    CLEAR
  END
  @ 4,0 TO 20,79 DOUBLE
  @ 2,0  SAY 'REBEL DISTRIBUTORS'
  @ 3,0  SAY 'Inventory Transmittal System'
  @ 3,72 SAY DATE()
RETURN


PROCEDURE NTX
  CLOSE DATABASES
  IF !FILE('RDINV.CDX')
    IF NETUSE("RDINV",.T.,5,"RDINV")
      @ 23,0 CLEAR
      @ 23,0 SAY 'Indexing Inventory File...'
      PACK
      GO TOP
      INDEX ON UPPER(RITEM)    TAG RDINV
      INDEX ON UPPER(BITEMNO)  TAG RDBITEM
      INDEX ON UPPER(MCITEMNO) TAG RDMCITEM
      USE
    ELSE
      MSGBOX("Can't re-index.  All other users must exit system.  Then try again.")
    ENDIF
  ENDIF

  CLOSE DATABASES
  NETUSE("RDINV",.F.,0,"RDINV")
  SET ORDER TO TAG RDINV
RETURN

FUNCTION INITIALSCREEN
  @ 0, 0 CLEAR
  WINDOW (3,10,13,70,10,40,2)
  WINDOW (4,12,12,68,10,40,1)
  COMPANY = "REBEL DISTRIBUTORS"
  PROGNAME = "Inventory Transmittal System"
  SPREAD(4," "+PROGNAME+" ",10,40)
  SPREAD(5,"1991-2002",10,40)
  *SPREAD(7,"by DAVID GRANT, INC. ",10,40)
  SPREAD(7,"Rebel Distributors Corp.",10,40)
  SPREAD(8,"Westlake Village, CA",10,40)
RETURN (.T.)

FUNCTION WINDOW
  PARAMETER TR,TC,BR,BC,SR,SC,LTYPE
  STORE 1 TO TRMOD, BRMOD
  STORE 2 TO TCMOD, BCMOD
  STORE .F. TO TRTEST,TCTEST,BRTEST,BCTEST
  STORE SR TO XBR,XTR
  STORE SC TO XBC,XTC
  IF SR=TR
    TRMOD = 0
    STORE .T. TO TRTEST
  ENDIF
  IF SR=BR
    BRMOD = 0
    STORE .T. TO BRTEST
  ENDIF
  IF SC=TC
    TCMOD = 0
    STORE .T. TO TCTEST
  ENDIF
  IF SC=BC
    BCMOD = 0
    STORE .T. TO BCTEST
  ENDIF
  DO WHILE .NOT. (TRTEST .AND. TCTEST .AND. BRTEST .AND. BCTEST)
    IF LTYPE = 1
      @ XTR,XTC TO XBR,XBC
    ELSE
      @ XTR,XTC TO XBR,XBC DOUBLE
    ENDIF
    @ XTR+1,XTC+1 CLEAR TO XBR-1,XBC-1
    XTR = XTR-TRMOD
    XTC = XTC-TCMOD
    XBR = XBR+BRMOD
    XBC = XBC+BCMOD
    IF XTR < TR
      XTR = TR
      TRTEST = .T.
    ENDIF
    IF XTC < TC
      XTC = TC
      TCTEST = .T.
    ENDIF
    IF XBR > BR
      XBR = BR
      BRTEST = .T.
    ENDIF
    IF XBC > BC
      XBC = BC
      BCTEST = .T.
    ENDIF
  ENDDO
  IF LTYPE = 1
    @ XTR,XTC TO XBR,XBC
  ELSE
    @ XTR,XTC TO XBR,XBC DOUBLE
  ENDIF
  @ XTR+1,XTC+1 CLEAR TO XBR-1,XBC-1
RETURN .T.


FUNCTION SPREAD
  PARAM ROW, MSG, DELAY, MID
  DELAY = IF(PCOUNT() < 3, 8, DELAY)
  MID = IF(PCOUNT() < 4, 40, MID)
  PRIVATE MLEN, MLOOP, MSG2, XX
  MSG2 = ''
  MLEN = LEN(MSG)
  FOR MLOOP = 1 TO INT(MLEN/2)
    MSG2 = SUBSTR(MSG, 1, MLOOP) + ;
    SUBSTR(MSG, MLEN + 1 - MLOOP, MLOOP)
    @ ROW, MID-MLOOP SAY MSG2
    FOR XX = 1 TO DELAY
    NEXT
  NEXT
  IF MLEN%2 = 1
    @ ROW, MID-MLOOP SAY MSG
  ENDIF
RETURN ('')


FUNCTION COPYFILE(MFILE)
  LOCAL DEFAULTDRIVE := "A",SCREEN := SAVESCREEN( 6,14,17,69 ),MENUTO := 1,;
        MDRIVE,MASK := .T.,SUCCESS := .F.,COLOR,SCOLOR, NUMFILE := .F.
  @ 6,14 CLEAR TO 17,69
  @ 7,15,16,68 BOX B_SINGLE+" "
  SCOLOR := SETCOLOR("G+/N")
  IF MFILE != "C:\COMMERCE\IMPORTS\WGTRAN.TDF"
    IF MFILE=="BUTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Burlington Drug?"
    ELSEIF MFILE="REB.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Morris & Dickson?"
    ELSEIF MFILE="AMTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Amerisource?"
    ELSEIF MFILE="REBEL.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to NC Mutual?"
    ELSEIF MFILE="xfiasitrb.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Longs?"
    ELSEIF MFILE="W2TRAN.TXT"
      @ 11,17 SAY "Prepare disk for shipment to legacy Walker File?"
    ELSEIF MFILE="AIVREB"
      @ 11,17 SAY "Prepare diskette for shipment to H.D. Smith?"
    ELSEIF MFILE="BWTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Bindley?"
    ELSEIF MFILE="ROTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Rochester?"
    ELSEIF MFILE="DAKOTA.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Dakota?"
    ELSEIF MFILE="rebwal.txt"
      @ 11,17 SAY "Prepare diskette for shipment to Walsh?"
    ELSEIF MFILE="DKTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to D & K Drug?"
    ELSEIF MFILE="MCTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to McKesson?"
    ELSEIF MFILE="KINRAY.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Kinray?"
    ELSEIF MFILE="CLTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Cardinal (AD)?"
    ELSEIF MFILE="CLTRANAL.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Cardinal (All)?"
    ELSEIF MFILE="DREBELA"
      @ 11,17 SAY "Prepare diskette for shipment to F. Dohmen?"
    ELSEIF MFILE="REBELSD.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Smith Drug?"
    ELSEIF MFILE="HTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Harvard Drug?"
    ELSEIF MFILE="BDTRAN.TXT"
      @ 11,17 SAY "Prepare diskette for shipment to Bellco Drug?"
    ELSEIF MFILE="DREBELA"
      @ 11,17 SAY "Prepare diskette for shipment to F. Dohman?"
    ENDIF
    SETCOLOR(SCOLOR)
    @ 13,35 PROMPT "Yes"
    @ 13,40 PROMPT "No "
    MENU TO MENUTO
    IF MENUTO == 1
      @ 11,17 SAY "                                                "
      @ 13,35 SAY "   "
      @ 13,40 SAY "   "
    
      @ 12,17 SAY "Select disk drive to prepare: " GET DEFAULTDRIVE PICT "!" VALID DEFAULTDRIVE $ "ABCF"
      READ
      IF MFILE=="BUTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\BUTRAN.TXT"
      ELSEIF MFILE="AMTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\AMTRAN.TXT"
      ELSEIF MFILE="REBEL.TXT"
        MDRIVE := DEFAULTDRIVE+":\REBEL.TXT"
      ELSEIF MFILE="XFIASITRB.TXT"
        MDRIVE := DEFAULTDRIVE+":\AMTRAN.TXT"
      ELSEIF MFILE="W2TRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\W2TRAN.TXT"
      ELSEIF MFILE="DKTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\DKTRAN.TXT"
      ELSEIF MFILE="FDTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\FDTRAN.TXT"
      ELSEIF MFILE="BWTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\BWTRAN.TXT"
      ELSEIF MFILE="ROTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\ROTRAN.TXT"
      ELSEIF MFILE="DAKOTA.TXT"
        MDRIVE := DEFAULTDRIVE+":\DAKOTA.TXT"
      ELSEIF MFILE="rebwal.txt"
        MDRIVE := DEFAULTDRIVE+":\BRTRAN.TXT"
      ELSEIF MFILE="BETRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\BETRAN.TXT"
      ELSEIF MFILE="MCTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\MCTRAN.TXT"
      ELSEIF MFILE="KINRAY.TXT"
        MDRIVE := DEFAULTDRIVE+":\KINRAY.TXT"
      ELSEIF MFILE="CLTRAN.TXT"
        * MDRIVE := DEFAULTDRIVE+":\CLTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\OFFERS.TXT"
      ELSEIF MFILE="CLTRANAL.TXT"
        * MDRIVE := DEFAULTDRIVE+":\CLTRANAL.TXT"
        MDRIVE := DEFAULTDRIVE+":\OFFERS.TXT"
        MFILE := "CARDINAL\OFFERS.TXT"
      ELSEIF MFILE="AIVREB"
        MDRIVE := DEFAULTDRIVE+":\AIVREB"
      ELSEIF MFILE="DREBELA"
        MDRIVE := DEFAULTDRIVE+":\DREBELA"
      ELSEIF MFILE="REBELSD.TXT"
        MDRIVE := DEFAULTDRIVE+":\REBELSD.TXT"
      ELSEIF MFILE="HTRAN.TXT"
        MFILE  := strHarvardVendorNumber+".DAT"
        MDRIVE := DEFAULTDRIVE+":\"+MFILE
        NUMFILE := .T.
      ELSEIF MFILE="BDTRAN.TXT"
        MDRIVE := DEFAULTDRIVE+":\BDTRAN.TXT"
      ENDIF

      MFILE := DISKDIR+"\EXPORT\"+MFILE

      IF LASTKEY() != 27
        IF DISKREADY( DEFAULTDRIVE+":" )
          COLOR := SETCOLOR( "W+*/N" )
          @ 12,17 SAY "     Preparing...  Please Wait...   "
          SETCOLOR( COLOR )
          IF NUMFILE
            *RENAME (DEFAULTDRIVE+":\X"+strHarvardVendorNumber+".DAT") TO (DEFAULTDRIVE+":\"+strHarvardVendorNumber)
            *RUN "XCOPY "+MFILE+" "+MDRIVE
            COPY FILE &MFILE TO &MDRIVE
          ELSE
            FILECOPY( MFILE,MDRIVE )
          ENDIF
          SUCCESS := .T.
        ELSE
          DO WHILE !DISKREADY( DEFAULTDRIVE+":" )
            TONE( 100,4 )
            @ 12,17 SAY "     Drive not ready.  Try again? " GET MASK PICT "Y"
            READ
            IF MASK .AND. LASTKEY() != 27
              IF DISKREADY( DEFAULTDRIVE+":" )
                COLOR := SETCOLOR( "W+*/N" )
                @ 12,17 SAY "     Preparing...  Please Wait...   "
                SETCOLOR( COLOR )
                IF NUMFILE
                  * RUN "XCOPY "+MFILE+" "+MDRIVE
                  COPY FILE &MFILE TO &MDRIVE
                  * RENAME (DEFAULTDRIVE+":\X"+strHarvardVendorNumber+".DAT") TO (DEFAULTDRIVE+":\"+strHarvardVendorNumber)
                ELSE
                  FILECOPY( MFILE,MDRIVE )
                ENDIF
                SUCCESS := .T.
              ENDIF
            ELSE
              EXIT
            ENDIF
          ENDDO
        ENDIF
      ENDIF
    ENDIF
  ENDIF
  SETCOLOR(SCOLOR)
  RESTSCREEN( 6,14,17,69,SCREEN )
RETURN SUCCESS

FUNCTION COPY_FILE()
  LOCAL DEFAULTDRIVE := "A",SCREEN := SAVESCREEN( 6,14,17,65 ),MENUTO := 1,;
        MDRIVE,MASK := .T.,SUCCESS := .F.,COLOR
  @ 6,14 CLEAR TO 17,65
  @ 7,15,16,64 BOX B_SINGLE+" "
  @ 11,17 SAY "Prepare diskette for shipment to Schein?"
  @ 13,35 PROMPT "Yes"
  @ 13,40 PROMPT "No "
  MENU TO MENUTO
  IF MENUTO == 1
    @ 11,17 SAY "                                             "
    @ 13,35 SAY "   "
    @ 13,40 SAY "   "

    @ 12,17 SAY "Select disk drive to prepare: " GET DEFAULTDRIVE PICT "!" VALID DEFAULTDRIVE $ "ABCF"
    READ
    MDRIVE := DEFAULTDRIVE+":\NTRAN.TXT"
    IF LASTKEY() != 27
      IF DISKREADY( DEFAULTDRIVE+":" )
        COLOR := SETCOLOR( "W+*/N" )
        @ 12,17 SAY "     Preparing...  Please Wait...   "
        SETCOLOR( COLOR )
        FILECOPY( "EXPORT\NTRAN.TXT",MDRIVE )
        SUCCESS := .T.
      ELSE
        DO WHILE !DISKREADY( DEFAULTDRIVE+":" )
          TONE( 100,4 )
          @ 12,17 SAY "     Drive not ready.  Try again? " GET MASK PICT "Y"
          READ
          IF MASK .AND. LASTKEY() != 27
            IF DISKREADY( DEFAULTDRIVE+":" )
              COLOR := SETCOLOR( "W+*/N" )
              @ 12,17 SAY "     Preparing...  Please Wait...   "
              SETCOLOR( COLOR )
              FILECOPY( "EXPORT\NTRAN.TXT",MDRIVE )
              SUCCESS := .T.
            ENDIF
          ELSE
            EXIT
          ENDIF
        ENDDO
      ENDIF
    ENDIF
  ENDIF
  RESTSCREEN( 6,14,17,65,SCREEN )
RETURN SUCCESS


FUNCTION GETMFGRANGE()
  LOCAL sSCREEN := SAVESCREEN(), sColor := SETCOLOR("W+/G")
  @ 10,23,16,56 BOX B_SINGLE+" "

  MMFGFROM := PADR(MMFGFROM,3)
  MMFGTO   := PADR(MMFGTO,3)

  @ 12,26 SAY "Enter range of manufacturers:"
  @ 14,38 SAY "to"
  @ 14,34 GET MMFGFROM PICT "!!!"
  @ 14,41 GET MMFGTO   PICT "!!!"
  READ
  MMFGFROM := ALLTRIM(MMFGFROM)
  MMFGTO   := ALLTRIM(MMFGTO)
  SETCOLOR( sColor )
  RESTSCREEN(,,,,sSCREEN)
RETURN NIL


FUNCTION WITHINRANGE()
  LOCAL lRET_VAL := .F., MMFGCODE := SUBSTR(RITEM,1,AT("-",RITEM)-1)
  IF MMFGCODE >= MMFGFROM .AND. MMFGCODE <= MMFGTO
    lRET_VAL := .T.
  ENDIF

  /* DEBUG STUFF
  SET DEVICE TO SCREEN
  @24,0 SAY PADR("MMFGCODE:"+MMFGCODE+" MMFGFROM:"+MMFGFROM+" MMFGTO:"+MMFGTO,79)
  INKEY(0)
  SET DEVICE TO PRINT
  */
RETURN lRET_VAL


FUNCTION MSGBOX( SSTRING )
  LOCAL SCOLOR := SETCOLOR( "W+/B" ),SSCREEN := SAVESCREEN(),;
        POS := (79-LEN(SSTRING))/2
  @ 10,POS-2,16,81-POS BOX B_DOUBLE+" "
  @ 12,POS SAY SSTRING
  @ 14,36 SAY " <OK> " COLOR "N/W"
  INKEY(0)
  SETCOLOR( SCOLOR )
  RESTSCREEN(,,,,SSCREEN)
RETURN NIL


FUNCTION StripAst( sString )
  DO WHILE AT("*",sString) > 0
    sString := STUFF(sString,at("*",sString),1," ")
  ENDDO
RETURN sString


FUNCTION CHECKORDUNIT( sString )
  LOCAL lRetVal := .T.
  IF ASCAN( {"EA","CT","CS","DZ","BX"},sString ) = 0
    TONE(100,1)
    MSGBOX("Valid entries are EA, BX, CS, CT, or DZ.  Please try again.")
    lRetVal := .F.
  ENDIF
RETURN lRetVal

FUNCTION CHECKREFTYPE( sString )
  LOCAL lRetVal := .T.
  IF ASCAN( {"R","I","S"},sString ) = 0
    TONE(100,1)
    MSGBOX("Valid entries are S, R, or I.  Please try again.")
    lRetVal := .F.
  ENDIF
RETURN lRetVal


FUNCTION STATUSBOX( SSTRING )
  LOCAL SCOLOR := SETCOLOR( "W+/B" ), POS := (79-LEN(SSTRING))/2
  _SSCREEN := SAVESCREEN()
  @ 10,POS-2,16,81-POS BOX B_DOUBLE+" "
  @ 13,POS SAY SSTRING COLOR "GR+*/B"
  SETCOLOR( SCOLOR )
RETURN NIL

FUNCTION CLEARSTATUS()
  RESTSCREEN(,,,,_SSCREEN)
RETURN NIL

