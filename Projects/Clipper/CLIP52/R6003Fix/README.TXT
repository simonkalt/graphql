The enclosed object file __WAIT_B.OBJ is used to fix a problem with
Clipper applications terminating at start up with either an "R6003
divide by zero" message or no error message at all. This problem has
recently reared it's head with the advent of new advanced design
microprocessors. These new design CPUs run "software timing loops" so
rapidly that a software timing loop module which is part of Nantucket
Tools II and CA-Tools III cannot work as designed. Not all Tools
functions use the software timing loop, but it appears to be linked into
your application if you're using either the CTUS.OBJ or CTUSP.OBJ
extended drivers. A few people have reported this problem without having
CTUS.OBJ or CTUSP.OBJ linked in (such as the MILLISEC() function).

The two processors currently exhibiting this condition are the AMD K5
and Cyrix 6x86. As of this writing, AMD apparently has no "fix" of their
own, while Cyrix has a "fix" posted on their web site at
http://www.cyrix.com  . Their fix consists of an executable named
PIPELOOP.EXE that is run from the AUTOEXEC.BAT file and causes timing
loops to be slowed down. Another "fix" is to turn off the internal
cache, which will significantly degrade all performance. Please note
this problem is not caused by a defect in the CPUs.

Just add __WAIT_B.OBJ to your link script ahead of the libraries and
ahead of the Tools IIIb extended driver file, if used.

PIPELOOP.EXE from Cyrix is no longer needed. It's been tested on Clipper
5.2 and 5.3 (see note below). It should work on 5.01a as well, but no
one has reported using it on that version. There's no need to maintain a
separate program version for Intel processors. It works fine on all
brands of CPUs. So far. <g>

The file name itself isn't significant. It's just a new name to
differentiate it from earlier versions.


5.3 USAGE
---------
CA has now released their own fix for this problem in the 5.3b patch in
the form of an object file named __WAIT_4.OBJ. I did some quick tests
and found it seems to work OK in 5.2 also, but as usual, use it at your
own risk.


COMPATIBILITY PROBLEMS
----------------------
As of June 1997, only one compatibility issue has come up. If you're
using the Tools IIIb function MILLISEC() and have a delay of less than
256 milliseconds, it no longer works.

I tested the 5.3b patch file __WAIT_4.OBJ in a 5.2e application and the
MILLISEC() function provides the same delays as without a patch file. If
you really need the MILLISEC() function in your 5.01a or 5.2e
application, you should try testing __WAIT_4.OBJ instead of
__WAIT_B.OBJ. The object file can be extracted from the 5.3b patch by
using the command  PATCH 53A_B /IGNOREMISSING  . This will create a
subdirectory named \OBJ and you'll find the object file in it.


HISTORY
-------
Several years ago, when the "fast" 486/66 CPU was released, this problem
started occurring among users of Nantucket Tools II. Someone on the
CompuServe ClipGer forum figured out the problem and released an object
file named __WAIT.OBJ. It worked perfectly.

When the same problem resurfaced with the AMD and Cyrix chips,
__WAIT.OBJ was tried and found to fix the problem when running in real
mode, but not protected mode. When this was discovered, a protected mode
version of __WAIT.OBJ was created by Ryszard Glab and posted on the
comp.lang.clipper newsgroup.

Subsequently it was found this protected mode __WAIT.OBJ module had
occasional GPF problems when used with specific nation module files and
the German language version. Malc Shedden of BlinkInc volunteered to
undertake the project to rewrite this module and the result is
__WAIT_B.OBJ, which is real, protected, and Blinker dual mode
compatible. Since it was released, it's been tested by dozens of persons
and it has fixed the problem 100% of the time, with only one
compatibility issue found (see above). However, keep in mind that it
certainly has not have been tested in all conceivable environments. If
you find a problem, please report it.

Thanks,

Ray Pesek 72270.650@compuserve.com
