****
**** EXPLODING WINDOW FUNCTION.
****

FUNCTION xplode( X1,Y1,X2,Y2,STR,SHADOW,HEADING )
PRIVATE ONE,TWO,THREE,FOUR,FRESTORE,SCOLOR,black,lightgray
External textpaint

black := 0
darkgray := 8
SCOLOR := SETCOLOR() 
SET CURSOR OFF
ONE := X1
TWO := X2
THREE := Y1
FOUR := Y2
DO WHILE ONE < TWO
	ONE += 1
	TWO -= 1
	THREE += 1
	FOUR -= 1
ENDDO

	@ ONE,THREE,TWO+1,FOUR BOX "        "

DO WHILE ONE > X1
	ONE -= 1
	TWO += 1
	THREE -= 1
	FOUR += 1
	@ ONE,THREE,TWO,FOUR BOX STR
	FOR I := 1 TO 25
	NEXT I
	@ ONE,THREE,TWO,FOUR BOX "        "
ENDDO
	@ ONE,THREE,TWO,FOUR BOX STR

	LENGTH := LEN(HEADING)
	RSPACE := LSPACE := INT(((FOUR - THREE -1) - LENGTH)/2)

	IF RSPACE + LSPACE + LENGTH != (FOUR - THREE - 1)
		LSPACE += 1
	ENDIF
IF Y1 > 0 .AND. X2 < 24 .AND. SHADOW
		textpaint(X1+1,Y1-1,X2,Y1-1,darkgray,black)
		textpaint(X2+1,Y1-1,X2+1,Y2-2,darkgray,black)
ENDIF
SETCOLOR(SCOLOR)
@ ONE,THREE + LSPACE+1 SAY HEADING
SET CURSOR ON
RETURN ""

*****
***** 3D BOX
*****

FUNCTION BOX3D ( A,B,C,D,E )
PRIVATE black,lightgray
black := 0
darkgray := 8
@ A,B,C,D BOX E
IF B > 0 .AND. C < 24
		textpaint(A+1,B-1,C,B-1,darkgray,black)
		textpaint(C+1,B-1,C+1,D-2,darkgray,black)
ENDIF
RETURN ""

*****
***** First letter capitalized
*****

Function FirstCap( Word )

If Len(Word) > 1
	Ret_Str := Upper(Substr(Word,1,1))+Lower(Substr(Word,2,Len(Word)-1))
Else
	Ret_Str := Word
Endif

Return Ret_Str


*****
***** Centering a Box
*****

Function Center_Box( Xchars,Ychars,Style )
Private Fromx,Fromy,Tox,Toy,Ret_Value

Fromx := 12-Int(Xchars/2)
Fromy := 39-Int(Ychars/2)
Tox := 12+Int(Xchars/2)
Toy := 39+Int(Ychars/2) + 1

@ Fromx,Fromy,Tox,Toy Box Style

Return .T.
