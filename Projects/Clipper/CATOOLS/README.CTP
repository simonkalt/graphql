CA-Clipper Tools 
Version 3.0
Protected-Mode Upgrade
December 1993

This file contains information about an upgrade to CA-Clipper Tools, 
Version 3.0, so that it can be used with CA-Clipper/ExoSpace to create 
applications that run in protected-mode.

You should have the following products installed before continuing:
  * CA-Clipper Tools, Version 3.0 
  * CA-Clipper, Version 5.2c or later
  * CA-Clipper/ExoSpace, Version 1.0a or later

Note that any versions of CA-Clipper/ExoSpace earlier than version 1.0e 
will be upgraded so that they are at the 1.0e release.  

Also note that it is highly recommended that you back up the current 
versions of the above products before performing this upgrade.

This file is divided into the following sections:
  * Installation Instructions
  * List of Installed Files
  * Linking CA-Clipper Tools Programs with CA-Clipper/ExoSpace 
  * The IPXCT Package
  * Special Considerations


=============================================================================
Installation Instructions
   
To perform this upgrade:

1. Insert the Installation diskette in drive A (or B).

2. Switch to the A (or B) drive.

3. Type INSTALL.

The Installer is invoked--this component manages the installation process 
from this point forward.  Follow the instructions on your screen to proceed 
with the installation; the CA-Clipper Tools and CA-Clipper/ExoSpace files 
will be installed in the locations you specify.  


=============================================================================
List of Installed Files

File            Contents
-----------------------------------------------------------------------------   
CTP.LIB         CA-Clipper Tools library that is required to link with 
                CA-Clipper/ExoSpace. 
                It is the protected-mode version of CT.LIB, and should be 
                used in the same way as CT.LIB.  

CTUSP.OBJ       CA-Clipper Tools object file that is required to link with 
                CA-Clipper/ExoSpace.   
                It is the protected-mode version of CTUS.OBJ, and should be 
                used in the same way as CTUS.OBJ. 

CTINTP.OBJ      CA-Clipper Tools object file that is required to link with 
                CA-Clipper/ExoSpace. 

                This file MUST BE USED in order to set the proper CA-Clipper
                Tools environment.  Failure to link in CTINTP.OBJ may cause 
                unpredictable results.

EXOSPACE.EXE    Updated CA-Clipper/ExoSpace program.

EXOSPACE.NG     Updated CA-Clipper/ExoSpace Norton Guide file.

NBDCHAT.PRG     Updated CA-Clipper Tools sample NETBIOS chatting program.


=============================================================================
Linking CA-Clipper Tools Programs with CA-Clipper/ExoSpace 

Getting your CA-Clipper Tools applications up and running with 
CA-Clipper/ExoSpace is really very simple.  You just need to change your
link command line and then relink your CA-Clipper Tools program using 
CA-Clipper/ExoSpace.  


Real-Mode Applications
----------------------

To create a real-mode Tools application, you typically link a CA-Clipper 
Tools program with a real-mode linker (be sure to include CT.LIB and, 
optionally, CTUS.OBJ).  

For example:

CLIPPER test 
RTLINK FILE test,ctus LIBRARY ct


Protected-Mode Applications
---------------------------

To create a protected-mode Tools application with CA-Clipper/ExoSpace, 
you need to:

1. Use the EXOSPACE linker instead of a real-mode linker (i.e., RTLINK).

2. Include CTINTP.OBJ (and optionally CTUSP.OBJ) in the FILE command.

3. Include CTP.LIB in the LIBRARY command.

4. Include the DOS25 package in the EXOSPACE PACKAGE command.

For example:

CLIPPER test 
EXOSPACE EXOSPACE PACKAGE DOS25 FILE test,ctusp,ctintp LIBRARY ctp


Once your CA-Clipper Tools application has been linked with 
CA-Clipper/ExoSpace, you should then be able to run the resulting .EXE as
easily as a CA-Clipper Tools real-mode application.  (You must run the 
application in a system which meets the minimum system requirements for 
CA-Clipper/ExoSpace.)

Please refer to your CA-Clipper/ExoSpace documentation for complete details
on linking, command syntax, and using script files.


=============================================================================
The IPXCT Package

The new version of EXOSPACE.EXE installed with this upgrade includes a new
package: IPXCT.  

IPXCT should only be used with protected-mode CA-Clipper Tools applications.

It allows you to utilize the point-to-point communications and 
print server functions.  It must be specified in the EXOSPACE PACKAGE script
command when linking an application that calls any point-to-point 
communication function (see Chapter 31) or Print Server Access function 
(see Chapter 27).

For example:

CLIPPER test
EXOSPACE EXOSPACE PACKAGE DOS25,IPXCT FILE test,ctusp,ctintp LIBRARY ctp


=============================================================================
Special Considerations

Communication functions may have problems if COM2 is used and you 
are not under DPMI (this may occur in the MINITERM sample program when 
using COM2).  In such a case, the package NOVM should be used.


                                  ***
