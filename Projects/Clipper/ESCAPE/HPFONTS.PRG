***
***  HPFONTS.PRG
***  Prints out a two column list of the font names supported by Escape.
***

clear 
@ 12,0 say "Indexing..."

use FONTS 
index on fontname to tempfonts
set index to tempfonts

****
* column positions
****
n_font1  = 24
n_pitch1 = 44
n_point1 = 57
n_font2  = n_font1 + 50
n_pitch2 = n_pitch1 + 50
n_point2 = n_point1 + 50

n_line   = 7              && line counter 
l_gray   = .F.            && specifies whether a gray bar should be printed.

n_lastln = round(reccount()/2,0)+17   && number of lines per column

@ 12,0 clear
@ 12,0 say "Printing..."

set console off
set print on
set device to print

****
* Select the internal 10 pitch Courier font, and letter size paper. 
* (3rd and 4th parameters default to Portrait orientation and 1 copy.)
*
* Include an error check of the return values to make sure HPSETUP()
* executes correctly.
****
hperr(HPSETUP("COURIER 10","I","LETTER"))

****
* print surrounding box and shadow
****
hperr(HPBOX(3,5,55,75))
HPSHADOW(3,5,55,75,.187,20)

****
* print title 
****
hperr(HPCENTER(HPBOLD("List of HP Laserjet Font Names"),n_line))
hperr(HPCENTER(HPBOLD("Supported by Escape"),n_line+1))

****
* print box and shadow surrounding title
****
hperr(HPBOX(5,20,10,61))
HPSHADOW(5,20,10,61,.187,20)

hperr(HPSETUP("LINEPRT 16","I"))     && change to 16 pitch line printer font

n_line = n_line + 9

****
* print column headings
****
@ n_line,n_font1 say "FONT NAME"
@ n_line,n_pitch1 say "PITCH"
@ n_line,n_point1 say "POINT"
@ n_line,n_font2 say "FONT NAME"
@ n_line,n_pitch2 say "PITCH"
@ n_line,n_point2 say "POINT"

****
* draw double lines at the top and bottom of the page
**** 
n_line = n_line + 1
hperr(HPBOX(n_line, n_font1-1, n_line, n_point2+5, 2))
hperr(HPBOX(n_lastln+1, n_font1-1, n_lastln+1, n_point2+5, 2))

hperr(HPATSAY(n_lastln+2, n_font1, "Table  1.0"))

****
* print fonts
****
for n_column = 1 to 2
  col = iif(n_column = 1,'1','2')
  n_line = 18
  l_gray = .F.

  do while n_line <= n_lastln .and. .not. eof()
    hperr(HPATSAY(n_line, n_font&col, fontname))
    hperr(HPATSAY(n_line, n_pitch&col, iif(empty(pitch),"Prop",pitch)))
    hperr(HPATSAY(n_line, n_point&col, point))
    ****
    * draw a gray bar every other line
    ****
    if l_gray 
      hperr(HPGRAYBAR(n_line, n_font1, n_point1+5 - n_font1, 0, 10))
      hperr(HPGRAYBAR(n_line, n_font2, n_point2+5 - n_font2, 0, 10))
    endif
    l_gray = (.not. l_gray)
    n_line = n_line + 1
    skip
  enddo
next       

hperr(HPRESET())
set print off
set console on
set device to screen  
close data
erase tempfonts.ntx
@ 12,0 clear
@ 12,0 say "Done..."
tone(2250,8)
tone(1950,10)
tone(2100,9)
tone(1500,20)
@ 20,0
return


