#include "sixcdx.ch"

FUNCTION NetUse( cDatabase,lOpenmode,nSeconds,cAlias )
  Local lForever
  lForever := (nSeconds == 0)

  DO WHILE (lForever .OR. nSeconds > 0)
    IF PCOUNT() == 4
      IF lOpenmode
        USE (cDatabase) ALIAS (cAlias) EXCLUSIVE 
      ELSE
        USE (cDatabase) ALIAS (cAlias) SHARED 
      ENDIF
    ELSE
      IF lOpenmode
        USE (cDatabase) EXCLUSIVE
      ELSE
        USE (cDatabase) SHARED
      ENDIF
    ENDIF

    IF !NETERR()
      RETURN (.T.)
    ENDIF
    INKEY(1)
    nSeconds = nSeconds-1
  ENDDO
RETURN (.F.)


FUNCTION FileLock( nSeconds )
  Local lForever
  IF FLOCK()
    RETURN (.T.)
  ENDIF
  lForever := (nSeconds==0)
  DO WHILE (lForever .OR. nSeconds > 0)
    INKEY(.5)
    nSeconds -= 0.5
    IF FLOCK()
      RETURN (.T.)
    ENDIF
  ENDDO
RETURN (.F.)


FUNCTION RecLock( nSeconds )
  Local lForever
  IF RLOCK()
    RETURN (.T.)
  ENDIF
  lForever := (nSeconds==0)
  DO WHILE (lForever .OR. nSeconds > 0)
    IF RLOCK()
      RETURN (.T.)
    ENDIF
    INKEY(.5)
    nSeconds -= 0.5
  ENDDO
RETURN (.F.)


FUNCTION AddRec( nWaitSeconds )
  Local lForever
  APPEND BLANK
  IF !NETERR()
    RETURN (.T.)
  ENDIF
  lForever := (nWaitSeconds==0)
  DO WHILE (lForever .OR. nWaitSeconds > 0)
    APPEND BLANK
    IF !NETERR()
      RETURN (.T.)
    ENDIF
    INKEY(.5)
    nWaitSeconds -= 0.5
  ENDDO
RETURN (.F.)


FUNCTION MustUse( sDatabase,lMode,nSeconds,sAlias )
  Local sScreen,sColor
  IF !NETUSE( sDatabase,lMode,nSeconds,sAlias )
    sColor := SetColor( "GR+/R" )
    sScreen := SaveScreen( 0,0,0,79 )
    @ 0,79-Len( "Waiting to open "+sDatabase+" data file..." ) SAY "Waiting to opening "+sDatabase+" data file..."
    DO WHILE !NETUSE( sDatabase,lMode,0,sAlias )
    ENDDO
    SetColor( sColor )
    RestScreen( 0,0,0,79,sScreen )
  ENDIF
RETURN (.T.)


