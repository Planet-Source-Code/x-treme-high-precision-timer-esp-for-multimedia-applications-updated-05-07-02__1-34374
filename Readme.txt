'*********************************************************************
' WARNING: ANY USE BY YOU IS AT YOUR OWN RISK. I provide this code
' "as is" without warranty of any kind, either express or implied,
' including but not limited to the implied warranties of
' merchantability and/or fitness for a particular purpose.
'*********************************************************************

This code is a VB implementation of the high precision multimedia timer which can be found
in the winmm.dll.

I wrote this code because of the same reason I created my multiple undo implementation:
I wanted to have a smart solution that can be easily added to different VB projects,
that doesn't clutter the code, that doesn't let depend my project on another custom
OCX or DLL file and that is object oriented as much as possible.

The demo project's intension is NOT to show that the multimedia timer is or is not more
accurate than the standard VB timer. It's primary task is just to show how to use the
timer class and how to instantiate and handle more than one timer object.

New in version 1.20:
--------------------
Timer moved from class to usercontrol, works quit similar to standard timer control now.
Shouldn't crash any longer in compiled EXE files.
Changed some methods and properties.
Added Event support.
Much easier handling.

New in version 1.01:
--------------------
High Precision Profiler Class added.


Usage:
------
If you want to use the timer in your project you just need to add the following files:
ctlHiPreTimer.ctl
ctlHiPreTimer.ctx	--> just copy into the same directory you place ctlHiPreTimer.ctl
modHiPreTimer.bas
PostMessageA.tlb	--> add as project reference

The High Precision Timer works similiar to the standard timer control. So have a look at
the VB documentation.

If you find this piece of code useful, then please vote for it at Planet-Source-Code.com:
http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34374&lngWId=1


Kind regards,
Sebastian Thomschke										  			05/08/2002
http://www.sebthom.de