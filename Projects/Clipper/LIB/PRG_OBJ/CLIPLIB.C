			/* Clipper functions in C */

/*		These are the functions that reside in
		this library. They are written in C for
		usage with Clipper:

----------------------------------------------------------

void	mix_it();         	       * randomizes
void	music(int);         	       * sound
void	hold_it(int);     	       * duration
void	stop_it();                     * nosound
int 	crand(int);       	       * returns random number
void	setcurs(int,int); 	       * sets size of cursor (0-13,0-13)
void textpaint(int,int,int,int,int,int)* changes color attribute of specified text rectangle

-----------------------------------------------------------
*/

/* # include <nandef.h> */
# include <d:\clipper5\include\extend.h>
# include <c:\tc\include\stdio.h>
# include <c:\tc\include\stdlib.h>
# include <c:\tc\include\time.h>
# include <c:\tc\include\dos.h>
# include <c:\tc\include\conio.h>
# include <c:\tc\include\mem.h>



/**** sound ****/

CLIPPER music()
{
	if (PCOUNT == 1)
		{
		sound(_parni(1));
		}
	else
		{
		_ret();
		}

	_ret();
}


/**** delay ****/

CLIPPER hold_it()
{
	if (PCOUNT == 1)
		{
		delay(_parni(1));
		}
	else
		{
		_ret();
		}

	_ret();
}


/**** nosound ****/

CLIPPER stop_it()
{
	if (PCOUNT == 0)
		{
		nosound();
		}
	else
		{
		_ret();
		}

	_ret();
}

/**** random ****/

int ret_val,rnd;

CLIPPER crand()
{
	if (PCOUNT == 1)
		{
		rnd = _parni(1);
		ret_val = random(rnd);
		_retni(ret_val);
		}
	else
		{
		_ret();
		}
}


/**** randomize ****/

CLIPPER mix_it()
{
	if (PCOUNT == 0)
		{
		randomize();
		}
	else
		{
		_ret();
		}

	_ret();
}

/* set cursor shape  ** not in use...

# define SET_CTYPE 1
# define VIDEO 0x10

CLIPPER setcurs()

{
	union REGS inregs;
	union REGS outregs;

	inregs.h.ah = SET_CTYPE;

	inregs.h.ch = _parni(1);
	inregs.h.cl = _parni(2);

	int86(VIDEO,&inregs,&outregs);

	_ret();
}
*/


/* change color attribute of specified text rectangle */

CLIPPER textpaint()

{
int x1,y1,x2,y2,attribute1,attribute2,offset,offset1,offset2,cntx,cnty;
unsigned a;
unsigned memloc = 0xb800;

x1=_parni(2);
y1=_parni(1);
x2=_parni(4);
y2=_parni(3);
attribute1=_parni(5);
attribute2=_parni(6);


  offset1 = 2*x1+y1*160+1;
  offset2 = 2*x2+y2*160+1;
  if ((offset1 > 0) && (offset1 <= 3999) && (offset2 > 0) && (offset2 <= 3999))
  {
    for(cnty = y1;cnty <= y2;cnty++)
    {
      for(cntx = x1;cntx <= x2;cntx++)
       {
	 offset = 2*cntx+cnty*160+1;
	 pokeb(memloc,offset,attribute1+(attribute2 <<4));
       }
    }
  }
_ret();
}
