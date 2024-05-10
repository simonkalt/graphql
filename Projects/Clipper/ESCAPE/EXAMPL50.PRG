************************************************************************
* HPSHADOW ( nBegRow, bBegCol, nEndRow, nEndCol, nWidth, nShade )
* 
* Description: Creates a shadow around a seperately drawn box that is
*              any width and any shade of gray.
* 
* Parameters:  nBegRow is the top row of box to shadow
*              nBegCol is the left column of box to shadow
*              nEndRow is the bottom row of box to shadow
*              nEndCol is the right column of box to shadow
*
*              nWidth  is the width of the shadow in inches
*              nShade  is the shade of gray to fill shadow with
*
* Example:     HPBOX(1, 1, 10, 20)
*              HPSHADOW(1, 1, 10, 20, (1/8), 60)  // 1/8" wide, 60% gray
************************************************************************
PROCEDURE hpshadow(nBegRow, nBegCol, nEndRow, nEndCol, nWidth, nShade)

* left shadow
hperr( HPIGRAYBOX( hpltoi(nBegRow)+nWidth, hpctoi(nBegCol)-nWidth,;
      hpltoi(nEndRow), hpctoi(nBegCol), 0, nShade))

* bottom shadow
hperr( HPIGRAYBOX( hpltoi(nEndRow), hpctoi(nBegCol)-nWidth,;
      hpltoi(nEndRow)+nWidth, hpctoi(nEndCol)-nWidth, 0, nShade))

RETURN


************************************************************************
* HPISHADOW ( nBegRow, bBegCol, nEndRow, nEndCol, nWidth, nShade )
* 
* Description: Creates a shadow around a seperately drawn box that is
*              any width and any shade of gray.
* 
* Parameters:  nBegRow is the top of box to shadow in inches
*              nBegCol is the left side of box to shadow in inches
*              nEndRow is the bottom of box to shadow in inches
*              nEndCol is the right side of box to shadow in inches
*
*              nWidth  is the width of the shadow in inches
*              nShade  is the shade of gray to fill shadow with
*
************************************************************************
PROCEDURE hpishadow(nBegRow, nBegCol, nEndRow, nEndCol, nWidth, nShade)


* left shadow
hperr( HPIGRAYBOX( nBegRow+nWidth, nBegCol-nWidth, ;
       nEndRow, nBegCol, 0, nShade))

* bottom shadow
hperr( HPIGRAYBOX( nEndRow, nBegCol-nWidth, ;
       nEndRow+nWidth, nEndCol-nWidth, 0, nShade))

RETURN


************************************************************************
* HPICOLCTR ( nBegRow, nBegCol, nWidth, cString )
*
* Description: Center a string over a column (useful for column
*              headings).
* 
* Parameters:  nBegRow is the distance for top of page to print string
*              nBegCol is where column starts in inches
*              nWidth  is width of column in inches
*              cString is the string to center over column
************************************************************************
PROCEDURE hpicolctr( nBegRow, nBegCol, nWidth, cString )

hperr( hpicenter( cString, nBegRow, nBegCol+( nWidth/2) ) )

RETURN


************************************************************************
* HPCOLCTR ( nBegRow, nBegCol, nWidth, cString )
*
* Description: Center a string over a column (useful for column
*              headings).
* 
* Parameters:  nBegRow is row to print string on
*              nBegCol is where column starts
*              nWidth  is width of column
*              cString is the string to center over column
************************************************************************
PROCEDURE hpcolctr( nBegRow, nBegcol, nWidth, cString)

hperr( hpcenter( cString, nBegRow, nBegCol+(nWidth/2)))

RETURN


************************************************************************
* HPGRAYON ()
*
* Description: Create and enable a macro for auto overlay that causes
*              the laser printer to print gray bars in simulation of
*              green bar paper.  This is useful in creating a utility that
*              turns this feature on for other, non-Clipper, applications
*              (ie, Lotus 123).
*
*              CAUTION:  Do not send a reset after setting up the laser 
*              printer for overlaying or the overlay will be disabled.
* 
* Parameters:  None.
*
************************************************************************
PROCEDURE hpgrayon
LOCAL nX,nY,lGray,nBarLength,nLPP

nBarLength := 8   // Setting for letter or legal size paper, portrait
                  // orientation. Change for other paper size/orientations.

nLPP := 60        // number of Lines Per Page, currently set for letter
                  // paper, portrait orientation.  Change for other paper
                  // size/orientations.

hperr( HPSTARTMAC(1) )
lGray := .F.
for nX := 0 to nLPP step 3
  for nY := 0 to 2
    if lGray
      hperr( HPIGRAYBAR(HPLTOI(nX+nY), 0, 8, 0, 10) )
    endif
  next
  lGray := (.not. lGray)
next
hperr( HPENDMAC(1) )
hperr( HPOVERLAY(1) )

RETURN


************************************************************************
* HPDECJUST ( cString, nRow, nCol )
*
* Description: To right justify a decimal number at its decimal point.
* 
* Parameters:  cString is the character representation of the decimal
*                      number to right justify
*              nRow    is the row to right justify the string.
*              nCol    is the column to position the decimal point on.
************************************************************************
PROCEDURE hpdecjust( cString, nRow, nCol)
LOCAL cLeft, cRight

cLeft := substr( cString, 1, at('.', cString))
cRight := substr( cString, at('.', cString)+1)

hperr( HPRJUST( cLeft, nRow, nCol))
hperr( HPATSAY( nRow, nCol, cRight))

RETURN


************************************************************************
* HPIDECJUST ( cString, nRow, nCol )
*
* Description: To right justify a decimal number at it's decimal point.
*              Position is specified in inches.
*              
* 
* Parameters:  cString is the character representation of the decimal
*                      number to right justify
*              nRow    is vertical position to print the string, in inches.
*              nCol    is the horizontal position, in inches, to right justify
*                      the decimal point.
************************************************************************
PROCEDURE hpidecjust( cString, nRow, nCol)
LOCAL cLeft,cRight

cLeft := substr(cString, 1, at('.', cString))
cRight := substr(cString, at('.', cString)+1)

hperr( HPIRJUST(cLeft, nRow, nCol) )
hperr( HPIATSAY(nRow, nCol, cRight) )

RETURN


************************************************************************
* hperr( nRet )
*
* Description: To conveniently handle error codes returned by Escape
* functions
*
* Parameters:   nRet is the return code value
************************************************************************

PROCEDURE hperr( nRet )
if nRet < 0
  set print off
  set console on
  set device to screen
  clear
  ?  "Escape error at line "+ltrim(str(procline(1)))
  ?? " in proc "+procname(1)
  ?  "Error #" + ltrim(str(nRet))
  do case
  case nRet == -1
    ? "Orientation of soft font and requested orientation don't match."
  case nRet == -2
    ? "File ESCAPE.FNT could not be found."
  case nRet == -3
    ? "Error reading ESCAPE.FNT, possibly corrupted."
  case nRet == -4
    ? "Inavlid font name."
  case nRet == -5
    ? "Soft font file could not be found."
  case nRet == -6
    ? "File error opening soft font.  FERROR() = "+ltrim(str(FERROR()))
  case nRet == -7
    ? "Invalid soft font format."
  case nRet == -8
    ? "Invalid # of lines per inch."
  case nRet == -9 .or. nRet == -10
    ? "Escape not initialized."
  case nRet == -11
    ? "Invalid paper specified."
  case nRet == -12
    ? "Wrong number of parameters passed."
  case nRet == -13
    ? "Column coordinates not valid for proportional fonts."
  case nRet == -14
    ? "Box coordinates reversed."
  case nRet == -15
    ? "Invalid feed method specified."
  case nRet == -16
    ? "Incorrect cartridge specified."
  otherwise
    ? "Internal error."
  endcase
  quit
endif
RETURN

