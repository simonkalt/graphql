select b
use rdinv
index on bwitemno to bwitem
set order to bwitem
select a
use c:\qryMatch
x = 0
do while .not. eof()
	select b
	seek a->ndc
	if .not. eof()
	  replace bwin with a->bwin
	endif
	select a
	skip
enddo

? "The end"