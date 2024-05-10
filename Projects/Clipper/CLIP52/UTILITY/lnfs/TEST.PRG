FIELD ID, FNAME, SNAME

/*ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
  ³Function ³Main                                                            ³
  ³Purpose  ³Well, guess what?  It's the main function                       ³
  ³Date-Time³02-07-97 05:37pm                                                ³
  ³Copyright³Dark Black Software Ltd.             Date updated: 02-07-97     ³
  ³Author   ³Ian 'DrDebug' Day                    Time updated: 05:37:57pm   ³
  ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ*/
FUNCTION Main

   LOCAL nHandle, aFiles, cFileName
   LOCAL aStru := {{"ID",    "C", 6, 0},;
                   {"FNAME", "C", 30, 0},;
                   {"SNAME", "C", 30, 0}}

   altd()

   aFiles := Directory("C:\", "V")
   IF Len(aFiles) > 0
      ? aFiles[1][1]
   ENDIF

   aFiles := Directory("C:\Program Files\*.*", "D")

   SortItOut(aFiles)

   IF LFNSupport()
      ? "Long filename support is available"
      cFileName := "Long Filename"
   ELSE
      ? "No Long filename support"
      cFileName := "SHORT"
   ENDIF

   IF !File(cFileName + ".DBF")
      dbCreate(cFileName, aStru)
   ENDIF

   ? "LFNShort('" + cFilename + "'): " + LFNShort(cFilename + ".DBF")

   USE (cFileName) NEW ALIAS Names

   INDEX ON ID TO (cFileName)

   SET INDEX TO (cFileName)

RETURN (NIL)

/*ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
  ³Function ³SortItOut                                                       ³
  ³Purpose  ³Sort out the Directory() array                                  ³
  ³Date-Time³02-07-97 05:37pm                                                ³
  ³Copyright³Dark Black Software Ltd.             Date updated: 02-07-97     ³
  ³Author   ³Ian 'DrDebug' Day                    Time updated: 05:37:37pm   ³
  ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ*/
FUNCTION SortItOut(aFiles)

   LOCAL nLoop

   FOR nLoop := Len(aFiles) TO 1 STEP -1
      aFiles[nLoop] := aFiles[nLoop][1]
   NEXT nLoop

RETURN (NIL)
