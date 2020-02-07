// Reminder.cpp : Defines the entry point for the application.
//

#include "stdafx.h"

int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
 	// TODO: Place code here.
	Private Sub cmdRun_Click()
	Call Remind
	End Sub


	Public Sub Remind()

Dim Response
Dim t1 As Double
Dim t2 As Double

On Error GoTo errhandler
t1 = 1625
t2 = 1830

    If CDbl(Mid(Now(), 11, 6)) >= t2 Then
        End
    End If
    
    If CDbl(Mid(Now(), 11, 6)) >= t1 Then 'Or CDbl(Mid(Now(), 11, 6)) = t2 Or CDbl(Mid(Now(), 11, 6)) = t3 Then
        Response = MsgBox("Open FTE_CoE_Tracking and Foglio Presenze?", vbYesNo + vbMsgBoxSetForeground + vbApplicationModal, "Reminder")
            If Response = vbYes Then
                End
            ElseIf Response = vbNo Then
                    Call wait(3600)  'wait 1 hour
            End If
    End If
    
Call wait(1)
errhandler:
    Call clearstack
End Sub

Public Sub wait(intime As Single)
   
   Dim cont   As Integer
   Dim t1, t2 As Variant
   
   On Error GoTo errhanlder
   cont = 0

   t1 = Timer
   DoEvents

   Do While cont = 0
      t2 = Timer
      
      ' check for midnight
      If t2 < 1 Then
         t1 = 0
      End If
      
      If (t2 > (t1 + intime)) Then
         cont = 1
      End If
   Loop
   DoEvents
   
Call Remind

errhandler:
Call clearstack
End Sub

Function clearstack()

End Function


	return 0;
}



