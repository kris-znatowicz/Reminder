Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module Module1
	Public Sub Remind()
		
        Dim Response As VariantType
		Dim t1 As Double
        Dim t2 As Double
        'Dim t3 As Double
        On Error GoTo Errorhandler
		
		t1 = 1625
        t2 = 1830
        't3 = 1825

        If CDbl(Mid(CStr(Now), 11, 6)) >= t2 Then
            End
        End If

        If CDbl(Mid(CStr(Now), 11, 6)) >= t1 Then 'Or CDbl(Mid(CStr(Now), 11, 6)) = t2 Or CDbl(Mid(CStr(Now), 11, 6)) = t3 Then
            'UPGRADE_WARNING: Impossibile risolvere la proprietà predefinita dell'oggetto Response. Fare clic qui per ulteriori informazioni: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
            Response = MsgBox("Open FTE_CoE_Tracking and Foglio Presenze?", MsgBoxStyle.YesNo + MsgBoxStyle.MsgBoxSetForeground + MsgBoxStyle.ApplicationModal, "Reminder")
            If Response = MsgBoxResult.Yes Then
                End
            ElseIf Response = MsgBoxResult.No Then
                Call wait(3600) 'wait 1 hour
            End If
        End If

        Call wait(1)
Errorhandler:
        'If System.StackOverflowException Then
        End
        'End If
	End Sub
	
	Public Sub wait(ByRef intime As Single)
		
		Dim cont As Short
        Dim t1, t2 As VariantType
        On Error GoTo Errorhandler
		cont = 0
		
		'UPGRADE_WARNING: Impossibile risolvere la proprietà predefinita dell'oggetto t1. Fare clic qui per ulteriori informazioni: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
		t1 = VB.Timer()
		System.Windows.Forms.Application.DoEvents()
		
		Do While cont = 0
			'UPGRADE_WARNING: Impossibile risolvere la proprietà predefinita dell'oggetto t2. Fare clic qui per ulteriori informazioni: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			t2 = VB.Timer()
			
			' check for midnight
			'UPGRADE_WARNING: Impossibile risolvere la proprietà predefinita dell'oggetto t2. Fare clic qui per ulteriori informazioni: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If t2 < 1 Then
				'UPGRADE_WARNING: Impossibile risolvere la proprietà predefinita dell'oggetto t1. Fare clic qui per ulteriori informazioni: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
				t1 = 0
			End If
			
			'UPGRADE_WARNING: Impossibile risolvere la proprietà predefinita dell'oggetto t1. Fare clic qui per ulteriori informazioni: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
			If (t2 > (t1 + intime)) Then
				cont = 1
			End If
		Loop 
		System.Windows.Forms.Application.DoEvents()
		
		Call Remind()
Errorhandler:
        End
	End Sub
End Module