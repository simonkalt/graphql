************************************************************************
* FUNCTION:     hperr( <expN> )
*
* DESCRIPTION:  To conveniently handle error codes returned by Escape
*               functions
*
* PARAMETERS:   <expN> is the return code value from previous Escape
*               function
*
* AUTHOR:       Peter M. Freese
*               1 June 1991
*               (c) 1991 Pinnacle Publishing, Inc.
************************************************************************

procedure hperr(rval)
if rval < 0
  set print off
  set console on
  clear screen
  ?  "Escape error at line "+ltrim(str(procline(1)))
  ?? " in proc "+procname(1)
  ?  "Error #" + ltrim(str(rval))
  do case
  case rval = -1
    ? "Orientation of soft font and requested orientation don't match."
  case rval = -2
    ? "File ESCAPE.FNT could not be found."
  case rval = -3
    ? "Error reading ESCAPE.FNT, possibly corrupted."
  case rval = -4
    ? "Inavlid font name."
  case rval = -5
    ? "Soft font file could not be found."
  case rval = -6
    ? "File error opening soft font.  FERROR() = "+ltrim(str(FERROR()))
  case rval = -7
    ? "Invalid soft font format."
  case rval = -8
    ? "Invalid # of lines per inch."
  case rval = -9 .or. rval = -10
    ? "Escape not initialized."
  case rval = -11
    ? "Invalid paper specified."
  case rval = -12
    ? "Wrong number of parameters passed."
  case rval = -13
    ? "Column coordinates not valid for proportional fonts."
  case rval = -14
    ? "Box coordinates reversed."
  case rval = -15
    ? "Invalid feed method specified."
  case rval = -16
    ? "Incorrect cartridge specified."
  otherwise
    ? "Internal error."
  endcase
  quit
endif
return


