#include "box.ch"
#include "set.ch"
#include "inkey.ch"
#include "directry.ch"

PRIVATE aScriptFiles := {}, aSorted := {}, aLogFiles := {}, sRunString := "",;
cErrorMemo := "", crlf := CHR(13)+CHR(10), sScreen, aLogFileNames := {},;
nHandle := 0, nPos := 0

PARAM sDatabaseName, sServerName, sUserName, sPassword

IF PCOUNT() < 3
  ??
  ? "USAGE: cappsscr <Database Name> <Server Name> <User Name> <Password>"
  ?
  QUIT
ELSEIF PCOUNT() == 3
  sPassword := ""
ENDIF

IF FILE("CAPPSSCR.LOG")
  RUN ("DEL CAPPSSCR.LOG")
ENDIF

nHandle := FCREATE("CAPPSSCR.LOG")
FWRITE(nHandle,("This is the log file for a CAPPS server upgrade ran "+DTOC(DATE())+" "+TIME()+"."+crlf+REPLICATE("-",73)+crlf),150)
FCLOSE(nHandle)

aScriptFiles := DIRECTORY("*.SQL")
cnt := 0

IF LEN(aScriptFiles) > 0

  aSorted := ASORT(aScriptFiles,,,{|x,y| VAL(SUBSTR(x[F_NAME],GetNumericPosition(x[F_NAME]),AT(".",x[F_NAME])-1)) < VAL(SUBSTR(y[F_NAME],GetNumericPosition(y[F_NAME]),AT(".",y[F_NAME])-1))})
  FOR i := 1 to LEN(aScriptFiles)
    AADD(aLogFiles,SUBSTR(aScriptFiles[i,F_NAME],1,LEN(aScriptFiles[i,F_NAME])-3)+"log")
    AADD(aLogFileNames,SUBSTR(aScriptFiles[i,F_NAME],1,LEN(aScriptFiles[i,F_NAME])-3)+"nam")
    sRunString := "OSQL -d "+sDatabaseName+" -S "+sServerName+" -U "+sUserName+" -P "+sPassword+" -i "+aScriptFiles[i,F_NAME]+" -o "+aLogFiles[i]
    ?
    ? "Preparing log file..."
    nHandle := FCREATE(aLogFileNames[i])
    FWRITE(nHandle,(crlf+(aLogFiles[i])+crlf+"------------"+crlf),28)
    FCLOSE(nHandle)
    set color to "W+/N"
    ? "Running script " + aScriptFiles[i,F_NAME] + "..."
    set color to "W/N"
    RUN (sRunString)
    ? "Merging log files..."
    ?
    RUN ("COPY CAPPSSCR.LOG+"+aLogFileNames[i]+"+"+aLogFiles[i]+" TMP.LOG")
    ? "Cleaning temp log files..."
    RUN ("DEL "+aLogFiles[i])
    ? "Cleaning more temp log files..."
    RUN ("DEL "+aLogFileNames[i])
    ? "Cleaning old log file..."
    RUN ("DEL CAPPSSCR.LOG")
    ? "Renaming temp log file..."
    RUN ("REN TMP.LOG CAPPSSCR.LOG")
  NEXT i
ELSE
  ? "No files found to process."
ENDIF

?
?
set color to "G+/N"
? "Done."
set color to "W/N"
? ""

@ 24,0 SAY ""


Function GetNumericPosition(sFileName)
  Local nLen := Len(sFileName), lNumeric := .F., sStr
  FOR i := 1 to nLen
    sStr := SUBSTR(sFileName,i,1)
    IF VAL(sStr) > 0 .OR. sStr = "0"
      lNumeric := .T.
      EXIT
    ENDIF
  NEXT i
  IF lNumeric
    nPos := i
  ELSE
    nPos := 0
  ENDIF
Return (nPos)

