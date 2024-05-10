/*
** $Header:   C:/sw/clpcom/vcs/include/slvrclip.chv   1.3   03 Apr 1991 10:35:20   davidn  $
*/


/*
** Clipper 5.0 Include file for the SilverWare SilverClip library.
**
** Copyright (C) 1991, SilverWare Inc.
*/

/*
** Constants for general use.
**
*/

#define SWTRUE		.T.		/* Logical .T. */
#define SWFALSE		.F.		/* Logical .F. */
#define SWSECOND	18		/* 1 second value */


/*
** Constants for COM ports.
**
*/


#define COM1		0		/* COM1 port value */
#define COM2		1		/* COM2 port value */
#define COM3		2		/* COM3 port value */
#define COM4		3		/* COM4 port value */
#define COM5		4		/* COM5 port value */
#define COM6		5		/* COM6 port value */
#define COM7		6		/* COM7 port value */
#define COM8		7		/* COM8 port value */
#define COM9		8		/* COM9 port value */
#define COM10		9		/* COM10 port value */
#define COM11		10		/* COM11 port value */
#define COM12		11		/* COM12 port value */
#define COM13		12		/* COM13 port value */
#define COM14		13		/* COM14 port value */
#define COM15		14		/* COM15 port value */
#define COM16		15		/* COM16 port value */
#define COM17		16		/* COM17 port value */
#define COM18		17		/* COM18 port value */
#define COM19		18		/* COM19 port value */
#define COM20		19		/* COM20 port value */
#define COM21		20		/* COM21 port value */
#define COM22		21		/* COM22 port value */
#define COM23		22		/* COM23 port value */
#define COM24		23		/* COM24 port value */
#define COM25		24		/* COM25 port value */
#define COM26		25		/* COM26 port value */
#define COM27		26		/* COM27 port value */
#define COM28		27		/* COM28 port value */
#define COM29		28		/* COM29 port value */
#define COM30		29		/* COM30 port value */
#define COM31		30		/* COM31 port value */
#define COM32		31		/* COM32 port value */
#define COM33		32		/* COM33 port value */
#define COM34		33		/* COM34 port value */


/*
** Constants for IRQs and Base I/O Addresses
**
*/

#define IRQ2		2		/* Hardware Interrupt Request #2 */
#define IRQ3		3		/* Hardware Interrupt Request #3 (COM2) */
#define IRQ4		4		/* IRQ #4 (COM1) */
#define IRQ5		5		/* IRQ #5 (Parallel Port 2) */
#define IRQ6		6		/* IRQ #6 (Diskette Controller) */
#define IRQ7		7		/* IRQ #7 (Parallel Port 1) */

/*
** Note: IRQ 8-15 are only available on AT class machines.  Only IRQ9 through
** IRQ 12 and IRQ 14 through 15 are brought to the buss.  Also note that
** IRQ9-IRQ12 and IRQ14-IRQ15 are only available on the small 36 pin connector
**
*/

#define IRQ8		8		/* IRQ #8 (Realtime Clock) */
#define IRQ9		9		/* IRQ #9 (IRQ 2 Redirection) */
#define IRQ10		10		/* IRQ #10 */
#define IRQ11		11		/* IRQ #11 */
#define IRQ12		12		/* IRQ #12 */
#define IRQ13		13		/* IRQ #13 (Coprocessor) */
#define IRQ14		14		/* IRQ #14 (Fixed Disk Controller) */
#define IRQ15		15		/* IRQ #15 */

#define COM1BASEIO	1016		/* Base Address of COM1 */
#define COM2BASEIO	760		/* Base Address of COM2 */


/*
** NOTE: The default BASE I/O Address and IRQ number
** for COM3 through COM8 are for the standard IBM Dual Async Adapter/A.
** The IBM Dual Async Adapter/A is a Micro-Channel Board.
*/

#define COM3BASEIO	12832		/* Base Address of COM3 */
#define COM4BASEIO	12840		/* Base Address of COM4 */
#define COM5BASEIO	16928		/* Base Address of COM5 */
#define COM6BASEIO	16936		/* Base Address of COM6 */
#define COM7BASEIO	21024		/* Base Address of COM7 */
#define COM8BASEIO	21032		/* Base Address of COM8 */



/*
** Values for usOptions parameter in SWCOpenComm() function
**
*/

#define SWNOASDTR	1		/* Do not Assert DTR */
#define SWNOASRTS	2		/* Do not Assert RTS */
#define SWCUSEINT14INTERFACE	4	/* Use Int 0x14 Interface */	

/*
** Values for usOptions parameter in SWCCloseComm() function
**
*/

#define SWASDTR		1		/* Assert DTR */
#define SWASRTS		2		/* Assert RTS */
#define SWDEASDTR	4		/* De-Assert DTR */
#define SWDEASRTS	8		/* De-Assert RTS */
#define SWNOFREE	16		/* Do not free() memory at close comm */

/*
** Error return codes
*/

#define SWCSUCCESSFUL		0	/* Successful */
#define SWCINVALIDPARAMETER	-500	/* An invalid parameter was specified */
#define SWCQUEUEISFULL		-499 	/* Attempted to insert a character to a full queue */
#define SWCQUEUEISEMPTY		-498 	/* Attempted to remove a character from an empty queue */
#define SWCINVALIDPORT 		-497 	/* An invalid port was specified when opening the port */
#define SWCNOUARTPRESENT	-496	/* No UART present at specified I/O Address */
#define SWCNOINTAVAILABLE	-495 	/* All hardware interrupts have been used */
#define SWCNOINTSPECIFIED	-494	/* Attempted to open a port and no interrupt was specified */
#define SWCNOMEMORY		-493 	/* Insufficient memory */
#define SWCTIMEDOUT		-492 	/* Timed out */
#define SWCABORT		-491	/* Function aborted */
#define SWCNOTFOUND		-490	/* Not Found */
#define SWCNOCARRIER		-489 	/* Carrier detect not asserted */
#define SWCUNABLETOOPENFILE	-488 	/* Unable to open specified file */
#define SWCBADBLOCKCHECK	-487 	/* Bad Block Check (CRC/CHECKSUM) */
#define SWCNOREMOTERESPONSE	-486	/* No Response from remote */
#define SWCLOCALABORT		-485 	/* Function aborted by local user */
#define SWCREMOTEABORT 		-484 	/* Function aborted by remote user  */
#define SWCBADBLOCKNUMBER	-483 	/* Block Number is bad */
#define SWCBLOCKFRAMINGERROR	-482 	/* Block Framing error */
#define SWCRECEIVEQUEUETOOSMALL	-481 	/* Receive buffer not big enough */
#define SWCEOTDETECTED 		-480 	/* EOT Detect during file receive */
#define SWCMAXLENGTHREACHED	-479 	/* Maximum length reached */
#define SWCTERMINATORREACHED	-478 	/* Terminator reached */
#define SWCQUEUEWENTEMPTY	-477 	/* Queue Went Empty */
#define SWCFEATURENOTAVAILABLE	-476	/* Feature Not Available */

/*
** Constants for various types of Parity.
**
*/

#define SWPARITYNONE		0	/* No Parity */
#define SWPARITYODD	      	8	/* Odd parity */
#define SWPARITYEVEN	      	24	/* Even parity */

#define SWPARITYSTICK	      	40	/* Parity always 1 */
#define	SWPARITYMARK		40	/* Parity always 1 */

#define SWPARITYMSTICK	      	56	/* Parity always 0 */
#define	SWPARITYSPACE		56	/* Parity always 0 */


/*
** Constants for bits returned by the SWTXBlcked() function
**
*/

#define TXWFORXON	1	/* Transmit Waiting for XON Character */
#define TXWFORCTS	2	/* Transmit Waiting for CTS to be asserted */
#define TXWFORDSR	4	/* Transmit Waiting for DSR to be asserted */
#define TXWFORCARRDET	8	/* Transmit Waiting for Carrier Detect */
#define TXWFORMANUAL	16	/* Transmit Waiting for Manual Enable */


/*
** Constants for bits returned by the SWGFloStat() function
**
*/

#define XONXOFFISENABLED	1	/* XON/XOFF Flow Control Enabled */
#define RTSCTSISENABLED 	2	/* RTS/CTS Flow Control Enabled */
#define DTRDSRISENABLED 	4	/* DTR/DSR Flow Control Enabled */
#define CARRDETISENABLED	8	/* Carrier Detec Flow Control Enabled */



/*
** Constants for modem on and off values.
**
*/

#define SWCSMOFF	"0"	/* Smartmodem OFF (initialization) */
#define SWCSMON 	"1"	/* Smartmodem ON (initialization) */
#define SWCSMMODEPULSE	0
#define SWCSMMODETONE	1

/*
** Constants for video attributes.
**
*/

#define SWBLACK 	0	/* Black */
#define SWBLUE		1	/* Blue */
#define	SWGREEN		2	/* Green */
#define	SWCYAN		3	/* Cyan */
#define	SWRED		4	/* Red */
#define	SWMAGENTA	5	/* Magenta */
#define	SWBROWN		6	/* Brown */
#define	SWWHITE		7	/* White */
#define	SWBLINK		128	/* Blinking */
#define SWINTENSE	8	/* High Intensity */


/*
** Defines constants used to access the POS registers
** in a MicroChannel Configuration.
**
*/

#define SWPOSREG0	0	/* POS Register #0 (Adapter I.D. Low Byte) */
#define SWPOSREG1	1	/* POS Register #1 (Adapter I.D. High Byte) */
#define SWPOSREG2	2	/* POS Register #2 (Option Data Byte #1) */
#define SWPOSREG3	3	/* POS Register #3 (Option Data Byte #2) */
#define SWPOSREG4	4	/* POS Register #4 (Option Data Byte #3) */
#define SWPOSREG5	5	/* POS Register #5 (Option Data Byte #4) */
#define SWPOSREG6	6	/* POS Register #6 Subaddress Extension (Low Byte)  */
#define SWPOSREG7	7	/* POS Register #7 Subaddress Extension (High Byte) */

#define SWPOSMINCHANNEL 0	/* Minimum Channel Number */
#define SWPOSMAXCHANNEL 7	/* Maximum Channel Number */

/*
** POS Adapter Identification Constant for the IBM Dual Async Adapter
*/

#define SWPOSIBMDUALASYNC	61183	/* IBM Dual Async Adapter */


/*
** POS Adapter Identification Constants for DigiBoard
*/

#define SWPOS16DIGIID	28644		/* Sixteen Channel Adapter */
#define SWPOS8DIGIID	28645		/* Eight Channel Adapter   */
#define SWPOS4DIGIID	28646		/* Four Channel Adapter    */


