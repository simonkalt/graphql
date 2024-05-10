use rdinv
c = 1
do while .not. eof()
	replace tempid with c
	skip
	c = c + 1
enddo