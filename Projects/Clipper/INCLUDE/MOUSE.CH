/*
Copyright(C) Delcom-Deltranik International Software Engineering 1990, 1991.

$owner: DAVEB$

$version: 1.2$ $date: March 15, 1993$ $time: 00:54:27 PM$

*/
static mouse_id_string := "$id: ssvcid mouse.ch 1.2 March 15, 1993 00:54:27 PM$"

/* $nokeywords$ */

/*
	This is the DOUBLE click timeout.  If a second click occurs with
	the timeout it will be deemed a double click
*/
#define	MOUSE_TIMEOUT		0.5

/*
	These define the offsets into the mouse event array.

	ex. 	IF (aMouseEvent[MOUSE_ROW] == MAXROW())
				...
			ENDIF
*/
#define	MOUSE_VALUE			1
#define	MOUSE_ROW			2
#define	MOUSE_COL			3
#define	MOUSE_CLICK_TIME	4
#define	MOUSE_TYPE			5

/*
	These define the type of mouse click that occured.

	ex. 	IF (aMouseEvent[MOUSE_VALUE] == MOUSE_SINGLE_CLICK)
				...
			ENDIF
*/
#define	MOUSE_SINGLE_CLICK	1
#define	MOUSE_DOUBLE_CLICK	2

/*
	Used with IsMouse functions ONLY.

	ex. 	DO WHILE (IsMouse(MOUSE_LEFT))
				...
			ENDDO
*/
#define	MOUSE_LEFT		1
#define	MOUSE_RIGHT		2
#define	MOUSE_BOTH		3
#define	MOUSE_MIDDLE	4
#define	MOUSE_HIDDEN	5

/*
	Used with SetMouse and IsMouse functions.

	ex. 	IF (IsMouse(MOUSE_OFF))
				SetMouse(MOUSE_ON)
			ENDIF
*/
#define	MOUSE_OFF				6
#define	MOUSE_ON				7
#define	MOUSE_WAIT			8
#define	MOUSE_INSTALL		9
#define	MOUSE_UNINSTALL	10

/*
	Used to set the type of mouse driver to use.
*/
#define MOUSE_SMOOTH		1
#define MOUSE_SMART			2
#define MOUSE_STANDARD	4
#define MOUSE_ANY				7

/*
	Makes sure all mouse functions are included in the link.
*/
EXTERN	ClearMouse
EXTERN	InitMouse
EXTERN	InMouse
EXTERN	IsMouse
EXTERN	MouseCol
EXTERN	MouseRow
EXTERN	MouseInCol
EXTERN	MouseInRow
EXTERN	ResetMouse
EXTERN	SetMouse
EXTERN	SetMousePos
EXTERN	SetMouseRange
