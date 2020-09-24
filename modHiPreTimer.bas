Attribute VB_Name = "modHiPreTimer"
'*********************************************************************
' High Precision Timer v1.20
' Copyright ©2002 by Sebastian Thomschke, All Rights Reserved.
' http://www.sebthom.de
'*********************************************************************
' If you like this code, please vote for it at Planet-Source-Code.com:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34374&lngWId=1
' Thank you
'*********************************************************************
' You are free to use this code within your own applications, but you
' are expressly forbidden from selling or otherwise distributing this
' source code without prior written consent.
'*********************************************************************
' WARNING: ANY USE BY YOU IS AT YOUR OWN RISK. I provide this code
' "as is" without warranty of any kind, either express or implied,
' including but not limited to the implied warranties of
' merchantability and/or fitness for a particular purpose.
'*********************************************************************
Option Explicit

Const WM_CHAR = &H102

'*********************************************************************
' Public Methods
'*********************************************************************
Public Sub TimerCallbackProc(ByVal TimerID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
   ' PostMessageA needs to be declared in a TypeLibrary (i.e. PostMessageA.tlb)
   ' or the application will crash when running compiled
   Call PostMessageA(dwUser, WM_CHAR, &HA, 0)
End Sub
