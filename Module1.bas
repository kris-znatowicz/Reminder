Attribute VB_Name = "Module1"
Option Explicit

Public EX1 As New Workbook

'myserver
 Public Declare Function RegMyServerObject Lib _
      "C:\Windows\syswow64\MyServerObject.DLL" _
      Alias "DllRegisterServer" () As Long
      
      
Public Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public TimerID As Long
Public TimerSeconds As Single

Public When As Double

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

 'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type


      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
   
      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
   
      Public nid As NOTIFYICONDATA




'Public RunWhen As Double
'Public Const cRunIntervalSeconds = 120 '3600    ' 1hour
'Public Const cRunWhat = "Remind"

Public Sub Remind()

Dim Response

'Dim EX2 As New Workbook
'Dim t1 As Double
'Dim t2 As Double
'Dim t3 As Double
'
'
't1 = 1625
't2 = 1725
't3 = 1825


If CDbl(Mid(Now(), 11, 6)) > 1030 And CDbl(Mid(Now(), 11, 6)) < 1040 Then
    Beep
    MsgBox "Coffee break."
End If

If CDbl(Mid(Now(), 11, 6)) > 1530 And CDbl(Mid(Now(), 11, 6)) < 1540 Then
    Beep
    MsgBox "Coffee break."
End If

If CDbl(Mid(Now(), 11, 6)) > 1230 And CDbl(Mid(Now(), 11, 6)) < 1430 Then
    Beep
    MsgBox "Lunch time."
End If

If CDbl(Mid(Now(), 11, 6)) >= 1625 Then ''Or CDbl(Mid(Now(), 11, 6)) >= t2 Or CDbl(Mid(Now(), 11, 6)) >= t3 Then
    Beep
    Response = MsgBox("Open Foglio Presenze?", vbYesNo + vbMsgBoxSetForeground + vbApplicationModal)
        If Response = vbYes Then
            MsgBox "Reminder will terminate & Tracking will open now. Manually open presenze please!"
'            VBAProject.Sheet1.Activate
'            FTE_CoE_Tracking.Sheet1.Visible = xlSheetVisible
            'Reminder.ThisWorkbook.Close
            'FTE_CoE_Tracking.Apri_Kris ''Open "X:\MEDICAL\R33W\Private\CoE_Utils\FTE_CoE_Tracking.xls" ''VBAProject.Apri_Kris
            
            EX1.Application.Visible = True
            EX1.Application.Workbooks.Open ("\\xfsvr02\servizi\PPN1.pro\Private\0_PPN1_Environments&Project Planning\PRESENZE PLANSOFT\2020_02_Feb_Presenze.xlsx")
            'EX2.Application.Visible = True
            'EX2.Application.Workbooks.Open ("P:\Z\ZNATOWIC\EXHIBITH\Presenze\Luglio.xls")
            
            Call EndTimer
            End
            'Reminder.ThisWorkbook.Close
        Else
            If Response = vbNo Then
               If CDbl(Mid(Now(), 11, 6)) >= 1725 Then ''Or CDbl(Mid(Now(), 11, 6)) >= t2 Or CDbl(Mid(Now(), 11, 6)) >= t3 Then
                    Beep
                    Response = MsgBox("Open Foglio Presenze?", vbYesNo + vbMsgBoxSetForeground + vbApplicationModal)
                    If Response = vbYes Then
                        MsgBox "Reminder will terminate & Tracking will open now. Manually open presenze please!"
                        'open excel files
                        EX1.Application.Visible = True
                        EX1.Application.Workbooks.Open ("\\xfsvr02\servizi\PPN1.pro\Private\Kris\Kris\plansoft\FTE_CoE_Tracking.xls")
                        'EX2.Application.Visible = True
                        'EX2.Application.Workbooks.Open ("P:\Z\ZNATOWIC\EXHIBITH\Presenze\Luglio.xls")
                        
                        Call EndTimer
                        End
                        'Reminder.ThisWorkbook.Close
                    Else
'                        If Response = vbNo Then
'                            If CDbl(Mid(Now(), 11, 6)) >= 1825 Then ''Or CDbl(Mid(Now(), 11, 6)) >= t2 Or CDbl(Mid(Now(), 11, 6)) >= t3 Then
'                                Beep
'                                Response = MsgBox("Open FTE_CoE_Tracking.xls and Foglio Presenze?", vbYesNo + vbMsgBoxSetForeground + vbApplicationModal)
'                                If Response = vbYes Then
'                                    Call EndTimer
'                                    End
'                                End If
'                            End If
'                        End If
                    End If
               End If
            End If
        End If
End If

If CDbl(Mid(Now(), 11, 6)) >= 1825 Then
    Beep
    MsgBox "Opening FTE_CoE_Tracking.xls. Manually open Foglio Presenze please. Reminder will terminate now. Vai a casa :o)"
    'open excel files
    EX1.Application.Visible = True
    EX1.Application.Workbooks.Open ("\\xfsvr02\servizi\PPN1.pro\Private\Kris\Kris\plansoft\FTE_CoE_Tracking.xls")
    'EX2.Application.Visible = True
    'EX2.Application.Workbooks.Open ("P:\Z\ZNATOWIC\EXHIBITH\Presenze\Luglio.xls")

    EndTimer
    End
End If
'StartTimer
End Sub

'Public Sub Wait(k As Integer)
'Dim i As Double
'    'waste time
'    i = Timer + k
'    Do While Timer < i
'    Loop
'    'Call Remind
'End Sub

'Sub MyPause(Optional ms As Long = 3000)
'    On Error Resume Next
'    Dim tc As Long
'    tc = GetTickCount
'    While GetTickCount < tc + ms: Sleep 1: DoEvents: Wend
'    Call Wait(1)
'End Sub

'Sub StartTimer()
'
'RunWhen = Now + TimeSerial(0, 0, cRunIntervalSeconds)
'Application.OnTime earliesttime:=RunWhen, Procedure:=cRunWhat, _
'     schedule:=True
'End Sub

'Sub StopTimer()
'   On Error Resume Next
'   Application.OnTime earliesttime:=RunWhen, _
'       Procedure:=cRunWhat, schedule:=False
'End Sub


'Sub StartTimer1()
'    TimerSeconds = 120 ' how often to "pop" the timer.
'    TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)
'End Sub
'
'Sub EndTimer1()
'    On Error Resume Next
'    KillTimer 0&, TimerID
'End Sub
Public Sub StartTimer()
    TimerSeconds = 3600 ' how often to "pop" the timer.
    TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)
End Sub

Public Sub EndTimer()
    On Error Resume Next
    KillTimer 0&, TimerID
End Sub
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal nIDEvent As Long, ByVal dwTimer As Long)
    '
    ' The procedure is called by Windows. Put your
    ' timer-related code here.
    '
    Call Remind
'    Call EndTimer
End Sub

'Sub Pause1h()
'
'      ' Pause for 1h
'      Application.OnTime When:=Now + TimeValue("01:00:00"), _
'         Name:="Remind"
'End Sub
'Sub Pause1m()
'
'      ' Pause for 1m
'      Application.OnTime When:=Now + TimeValue("00:01:00"), _
'         Name:="Remind"
'End Sub

