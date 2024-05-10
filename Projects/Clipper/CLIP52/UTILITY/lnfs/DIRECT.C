#include "clipper.api"

extern void *memcpy( void *__s1, const void *__s2, unsigned __n);
#pragma intrinsic (memcpy,strlen)

static void near TimeSub (char * cPtr, short nValue);
static ITEM near FillShort (struct find_t * Find);
static ITEM near FillLong (struct lfnfind_t * lfnFind);
static int  near FillAttributes (char * cBuffer, ULONG nAttrib);

typedef struct _MyVOLUMELABEL { /* vol */
   ULONG ulSerialNumber;
   char  cch;
   char  szVolLabel[12];
   ULONG ulDunno;
} MyVOLUMELABEL, * MyVOLUMELABELP;


/*ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
  ³Function ³Directory ()                                                     ³
  ³Purpose  ³Get directory (modified for OS/2 and long file names)            ³
  ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ*/
CLIPPER Directory (void)
{
   USHORT nCount = 1;
   int nErrCode;
   USHORT nSearchHandle;
   ITEM aRet;
   short nElements = 0;
   ITEM aFilespec = _param (1, VALUE_CHARACTER);
   ITEM aAttrib = _param (2, VALUE_CHARACTER);
   unsigned nAttrib = FILE_ARCHIVE | FILE_READONLY;
   char * cFilespec;
   int lGetVolume = FALSE;
   char * cRoot = "\\*.*";
   struct lfnfind_t lfnFind;
   struct find_t Find;

   if (aAttrib) {
      char * cAttrib = _VSTR (aAttrib);

      if (cAttrib != NULL) {
         char * cPtr = cAttrib;

         while (*cPtr) {
            char cByte = *cPtr;

            switch (cByte) {
               case 'H':   nAttrib |= FILE_HIDDEN;    break;
               case 'S':   nAttrib |= FILE_SYSTEM;    break;
               case 'D':   nAttrib |= FILE_DIRECTORY; break;
               case 'V':   lGetVolume = TRUE;         break;
            }

            cPtr++;
         }
      }
   }

   if (aFilespec) {
      cFilespec = _VSTR (aFilespec);
   } else {
      cFilespec = cRoot + 1;     // "*.*" We already have it in 'cRoot'
   }

   if (lGetVolume) {
      int nDrive;
      MyVOLUMELABEL vl;
      ITEM aEntry;

      if (cFilespec[1] == ':') {
         strcpy (cFilespec + 2, cRoot);
      } else {
         cFilespec = cRoot;
      }

      nErrCode = _f_first (cFilespec, &Find, FILE_VOLUME);

      if (nErrCode != 1) {
         Find.name[0] = '\0';
      }

      aEntry = FillShort (&Find);

      _ARRAYNEW (1);
      aRet = _GetGrip (_eval);

      _cAtPut (aRet, 1, aEntry);     /* Put item into array */
      _DropGrip (aEntry);

   } else {
      _ARRAYNEW (0);
      aRet = _GetGrip (_eval);

      if (_tlfn) {
         nErrCode = nSearchHandle = _f_firstlfn (cFilespec, &lfnFind, nAttrib);
      } else {
         nErrCode = _f_first (cFilespec, &Find, nAttrib);
      }

      while (nErrCode != 0) {
         ITEM aEntry;

         if (_tlfn) {
            aEntry = FillLong (&lfnFind);
         } else {
            aEntry = FillShort (&Find);
         }

         _cResize (aRet, 1);
         nElements++;
         _cAtPut (aRet, nElements, aEntry);     /* Put item into array */
         _DropGrip (aEntry);

         if (_tlfn) {
            nErrCode = _f_nextlfn (nSearchHandle, &lfnFind);
         } else {
            nErrCode = _f_next (&Find);
         }
      }

      if (_tlfn) {
         _f_closelfn (nSearchHandle);
      }
   }

   memcpy (_eval, aRet, 14);

   _DropGrip (aRet);
}

/*ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
  ³Function ³TimeSub ()                                                       ³
  ³Purpose  ³Do part of the time                                              ³
  ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ*/
static void near TimeSub (char * cPtr, short nValue)
{
   cPtr[0] = nValue / 10 + '0';
   cPtr[1] = nValue % 10 + '0';
}

/*ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
  ³Function ³FillShort ()                                                     ³
  ³Purpose  ³Fill short filename array entry                                  ³
  ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ*/
static ITEM near FillShort (struct find_t * Find)
{
   char cBuffer[10];
   ITEM aEntry;

   _ARRAYNEW (5);
   aEntry = _GetGrip (_eval);

   // 1 - Filename
   _cAtPutStr (aEntry, 1, Find->name, strlen (Find->name));

   // 2 - Size
   _putln (Find->size);
   _cAtPut (aEntry, 2, _tos);
   _tos--;

   // 3 - Date
   _putln (_dDMYToDate (Find->wr_date.nDay, Find->wr_date.nMonth, Find->wr_date.nYear + 1980));
   _tos->nType = VALUE_DATE;
   _cAtPut (aEntry, 3, _tos);
   _tos--;

   // 4 - Time
   TimeSub (cBuffer,     Find->wr_time.nHour);
   TimeSub (cBuffer + 3, Find->wr_time.nMinute);
   TimeSub (cBuffer + 6, Find->wr_time.nSecond);
   cBuffer[2] = ':';
   cBuffer[5] = ':';
   cBuffer[8] = '\0';
   _cAtPutStr (aEntry, 4, cBuffer, 8);

   // 5 - Attributes
   _cAtPutStr (aEntry, 5, cBuffer, FillAttributes (cBuffer, Find->attrib));

   return (aEntry);
}

/*ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
  ³Function ³FillLong ()                                                      ³
  ³Purpose  ³Fill long filename array entry                                   ³
  ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ*/
static ITEM near FillLong (struct lfnfind_t * lfnFind)
{
   char cBuffer[10];
   ITEM aEntry;

   _ARRAYNEW (5);
   aEntry = _GetGrip (_eval);

   // 1 - Filename
   _cAtPutStr (aEntry, 1, lfnFind->name, strlen (lfnFind->name));

   // 2 - Size
   _putln (lfnFind->size);
   _cAtPut (aEntry, 2, _tos);
   _tos--;

   // 3 - Date
   _putln (_dDMYToDate (lfnFind->wr_date.nDay, lfnFind->wr_date.nMonth, lfnFind->wr_date.nYear + 1980));
   _tos->nType = VALUE_DATE;
   _cAtPut (aEntry, 3, _tos);
   _tos--;

   // 4 - Time
   TimeSub (cBuffer,     lfnFind->wr_time.nHour);
   TimeSub (cBuffer + 3, lfnFind->wr_time.nMinute);
   TimeSub (cBuffer + 6, lfnFind->wr_time.nSecond);
   cBuffer[2] = ':';
   cBuffer[5] = ':';
   cBuffer[8] = '\0';
   _cAtPutStr (aEntry, 4, cBuffer, 8);

   // 5 - Attributes
   _cAtPutStr (aEntry, 5, cBuffer, FillAttributes (cBuffer, lfnFind->attrib));

   return (aEntry);
}

/*ÚÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
  ³Function ³FillAttributes ()                                                ³
  ³Purpose  ³Fill attrib buffer                                               ³
  ÀÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ*/
static int near FillAttributes (char * cBuffer, ULONG nAttrib)
{
   char * cPtr = cBuffer;

   if (nAttrib & FILE_READONLY) {
      *(cPtr++) = 'R';
   }

   if (nAttrib & FILE_HIDDEN) {
      *(cPtr++) = 'H';
   }

   if (nAttrib & FILE_SYSTEM) {
      *(cPtr++) = 'S';
   }

   if (nAttrib & FILE_DIRECTORY) {
      *(cPtr++) = 'D';
   }

   if (nAttrib & FILE_ARCHIVE) {
      *(cPtr++) = 'A';
   }

   if (nAttrib & FILE_VOLUME) {
      *(cPtr++) = 'V';
   }

   *cPtr = '\0';

   return (cPtr - cBuffer);
}
