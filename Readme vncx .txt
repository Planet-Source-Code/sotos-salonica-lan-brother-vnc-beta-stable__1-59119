VNCX Release 0.9.2.8 [12 November 2001]
------------------------------------

Author: Thong Nguyen (thongnguyen@mail.com)

Home: http://tummy.veridicus.com/tummy/programming/vncx


Introduction
------------

VNCX is a VNC client implemented as an ActiveX control and is intended for us by application developers.

VNC allows you to control desktops remotely and is very platform independent.  It is mainted by the AT&T research group.
You can find out more about VNC at:

http://www.uk.research.att.com/vnc/


Installation:
-------------

If you have any previous version of VNCX installed, deregister it with:

  REGSVR32.EXE VNCX.DLL

Register the VNCX.DLL from the command prompt by typing:

  REGSVR32.EXE VNCX.DLL

You must be in the directory containing VNCX.DLL, and REGSVR32.EXE must be in your PATH.  REGSVR32.EXE is usually located in C:\Windows\System or C:\WINNT\System32.


Documentation:
--------------

Documentation isn't included in this distribution but can be found at the following address:

http://tummy.veridicus.com/tummy/programming/vncx/documentation.asp

LICENSE:
--------

VNCX is pretty much free.  You must however agree to the following terms of use:

- You can't hold the author (me) liable for anything that may occur because of VNCX.

- If you plan to make money selling a product that is based on VNCX, you need to ask and get my permission to do so.  If you're unsure, email me and ask.

- Credit to the author (me) should be given in any software that uses VNCX.  You should do at least the following things:
   
	Include the following in your program documentation:

		- A small explanation of how your software uses VNCX.
		- VNCX's homepage (http://tummy.veridicus.com/tummy/programming/vncx)
		- The author (me :))

	Include the following in your software:
	
		- An easily accessible menu that displays the VNCX about box.
		  This can be done by calling the VNCViewer.AboutBox() method.
     
