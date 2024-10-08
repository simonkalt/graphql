FIELD ID, FNAME, SNAME

/*旼컴컴컴컴쩡컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴커
  쿑unction 쿘ain                                                            �
  쿛urpose  쿥ell, guess what?  It's the main function                       �
  쿏ate-Time�02-07-97 05:37pm                                                �
  쿎opyright쿏ark Black Software Ltd.             Date updated: 02-07-97     �
  쿌uthor   쿔an 'DrDebug' Day                    Time updated: 05:37:57pm   �
  읕컴컴컴컴좔컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴켸*/
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

/*旼컴컴컴컴쩡컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴커
  쿑unction 쿞ortItOut                                                       �
  쿛urpose  쿞ort out the Directory() array                                  �
  쿏ate-Time�02-07-97 05:37pm                                                �
  쿎opyright쿏ark Black Software Ltd.             Date updated: 02-07-97     �
  쿌uthor   쿔an 'DrDebug' Day                    Time updated: 05:37:37pm   �
  읕컴컴컴컴좔컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴켸*/
FUNCTION SortItOut(aFiles)

   LOCAL nLoop

   FOR nLoop := Len(aFiles) TO 1 STEP -1
      aFiles[nLoop] := aFiles[nLoop][1]
   NEXT nLoop

RETURN (NIL)
