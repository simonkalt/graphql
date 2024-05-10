***
*** HPCARTS.PRG :
*** Prints out a list of the HP cartridge fonts using
*** the Escape library.
***

clear
@ 12,0 say "Indexing..."

sele 0
use fonts alias fonts
index on fontname to tempfont
set index to tempfont

sele 0
use cartcont alias cartcont
index on cartletter+fontname+orient to tempcont
set index to tempcont

sele 0
use cartrdge alias cartrdge
index on cartletter to tempcart
set index to tempcart

****
* column positions
****              
n_cart   = 3
n_font   = 26
n_pitch  = 43
n_point  = 54
n_orient = 65
n_page   = 69

n_pg     = 0      && page number
n_lastln = 57     && maximum number of lines per page
n_line   = 58     && line counter
l_GRAY   = .F.    && specifies whether a GRAY bar should be printed

@ 12,0 clear
@ 12,0 say "Printing..."

set console off
set print on
set device to print

****
* Select the internal 10 pitch Courier font, and letter size paper.
* (3rd and 4th parameters default to Portrait orientation and 1 copy.
*
* Include an error check of the return values to make sure HPSETUP()
* executes correctly.
****
hperr(hpSETUP("COURIER 10","I","LETTER"))

do while .not. eof()

  if n_line > n_lastln
    n_line = 1
    n_pg = n_pg + 1

    ****
    * print page number and title
    ****                         
    @ n_line + 1,n_page say "Page " + str(n_pg,2)
    hperr(hpCENTER(HPBOLD("Table of HP Laserjet"), n_line))
    hperr(hpCENTER(HPBOLD("Cartridge Fonts"), n_line + 1))

    n_line = n_line + 5

    ****
    * print column headings
    ****                   
    @ n_line,n_font say "Font Name"
    @ n_line,n_pitch say "Pitch"
    @ n_line,n_point say "Point"
    @ n_line,n_orient say "Orientation"

    ****
    * draw double lines at top and bottom of page
    ****                                         
    n_line = n_line + 1
    hperr(hpBOX(n_line, n_cart-1, n_line, n_page+7, 2))
    hperr(hpBOX(n_lastln+1, n_cart-1, n_lastln+1, n_page+7, 2))
  endif

  ****
  * check if there is room on the page to print data for the next cartridge.
  ****
  if .not. isroom()
    loop
  endif

  ****
  * print dotted underline
  ****                    
  hperr(hpBOX(n_line, n_cart-1, n_line, n_page+7, 3))
  n_line = n_line + 1

  ****
  * print cartridge name
  ****                  
  @ n_line, n_cart say HPBOLD("Cartridge " + cartletter)
  @ n_line + 1, n_cart say cartname
  @ n_line + 2, n_cart say cartnumber

  ****
  *  print cartridge contents
  ****                       
  sele cartcont
  go top
  seek cartrdge-> cartletter

  do while cartletter = cartrdge->cartletter
    sele fonts
    seek cartcont->fontname
    hperr(hpATSAY(n_line, n_font, fontname))
    hperr(hpATSAY(n_line, n_pitch, iif(empty(pitch), "Prop", pitch)))
    hperr(hpATSAY(n_line, n_point, point))
    hperr(hpATSAY(n_line, n_orient, iif(cartcont->orient="P","Portrait","Landscape")))

    ****
    * draw a GRAY bar every other line
    ****                              
    if l_GRAY
      hperr(hpGRAYBAR(n_line, n_font, n_page+7-n_font, 0, 10))
    endif
    l_GRAY = (.not. l_GRAY)

    sele cartcont
    n_line = n_line + 1
    skip
  enddo

  l_GRAY = .F.            && no GRAY shading on first line
  n_line = n_line + 1

  sele cartrdge
  skip
enddo

hperr(hpRESET())
set print off
set device to screen
set console on       
close data
erase tempfont.ntx
erase tempcont.ntx
erase tempcart.ntx
@ 12,0 clear
@ 12,0 say "Done..."
tone(2250,8)
tone(1950,10)
tone(2100,9)
tone(1500,20)
@ 20,0
return

* /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ *

****
* Before printing cartridge contents:
* 
* Count the number of fonts that are contained in the cartridge.
* If there is enough room to fit it all on the same page, then
* print, otherwise begin the next page.
****
function isroom

n_currarea = select()

sele CARTCONT
seek CARTRDGE->cartletter

n_lines = 0
do while cartletter = cartrdge->cartletter
  n_lines = n_lines + 1
  skip
enddo

select(n_currarea)

if n_line + n_lines <=n_lastln
  return(.T.)
else
  n_line = n_lastln + 1
  return(.F.)
endif

return(.F.)
