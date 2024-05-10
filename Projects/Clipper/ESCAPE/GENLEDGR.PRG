***
*** GENLEDGR.PRG :
*** Prints out general ledger paper with the Escape box drawing
*** functions.
***

clear
@ 12,0 say 'Printing...'
set print on
set console off
set device to print

****
* Select 16 pitch line printer font, letter-size paper, portrait, 1 copy
* 4 lines per inch to get lines spaced correctly
****
hperr(hpsetup('LINEPRT 16', 'I', 'LETTER','P',1))
hperr(hpsetlpi(4))

hperr(hpbox(2,6,40,132,2))                  && outer box

for x = 3 to 39
  hperr(hpbox(x,6,x,132,iif(x=5,2,1)))      && horizontal lines
next

hperr(hpbox(2,13,40,13,1))                  && first single line
hperr(hpbox(2,19,40,19,2))                  && first vertical double line

n_col = 52
for y = 1 to 4
  hperr(hpbox(2,n_col,40,n_col,2))          && first vertical double line
  hperr(hpgraybox(2,n_col,40,n_col+4,1,10)) && gray 2 char box
  hperr(hpbox(2,n_col+6,40,n_col+8,1))      && single line box
  hperr(hpgraybox(2,n_col+10,40,n_col+16,1,10))
  hperr(hpbox(2,n_col+12,40,n_col+14,1))    && single line box
  n_col = n_col + 20
next

****
* Switch to 10 pitch for title
****
hperr(hpsetup("COURIER 10", "I"))

@ 0,5 SAY hpbold('PINNACLE PUBLISHING') + ' SAMPLE GENERAL LEDGER SHEET - '

hperr(hpsetup("COURIER 10I", "I"))
@ prow(),pcol() say 'Produced by Escape'

****
* Switch back to Line Printer for line numbers
****
hperr(hpsetup("LINEPRT 16", "I"))

for x = 6 to 39
  @ x,3 say str(x-5,2,0)
next

hperr(hpreset())
set device to screen
set print off
set console on
@ 12,0 clear
@ 12,0 say "Done..."
tone(2250,8)
tone(1950,10)
tone(2100,9)
tone(1500,20)
@ 20,0
return
