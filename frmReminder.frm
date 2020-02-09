VERSION 5.00
Begin VB.Form frmReminder 
   Caption         =   "Reminder"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4170
   Icon            =   "frmReminder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Exit Reminder"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuGeoSol 
         Caption         =   "&GeoSolutions on the Web"
      End
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdStop_Click()
Dim Response
         
Response = MsgBox("Are you sure you want to exit?", vbYesNo)

         
    If Response = vbYes Then
        Call EndTimer
        'Application.DisplayAlerts = False
        End 'Reminder.ThisWorkbook.Close
    End If
        
End Sub



Private Sub Command1_Click()
 Dim MyObj As New MyObject
         MyObj.MyProperty = "Hello"
         MsgBox MyObj.MyProperty
End Sub

Private Sub Form_Load()
StartTimer
Remind
'myserver
'On Error GoTo Err_DLL_Not_Registered
         Dim RegMyDLLAttempted As Boolean
         Dim MyObj As New MyServerObject.MyObject

         'The following statement will fail at run-time
         'if MyServerObject is not registered.
         MyObj.MyProperty = "Hello"
         Set MyObj = Nothing
         Exit Sub

'the form must be fully visible before calling Shell_NotifyIcon
       Me.Show
       Me.Refresh
       With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Reminder" & vbNullChar
       End With
       Shell_NotifyIcon NIM_ADD, nid

StartTimer

If (Mid(Now(), 11, 6)) > "7:30" And (Mid(Now(), 11, 6)) < "8:00" Then
    Beep
    MsgBox "Sei in ritardo! Dovresti ricuperare il tempo stasera."
End If
If (Mid(Now(), 11, 6)) > "8:30" And (Mid(Now(), 11, 6)) < "9:00" Then
    Beep
    MsgBox "Sei in ritardo! Dovresti ricuperare il tempo stasera."
End If
If (Mid(Now(), 11, 6)) > "9:30" And (Mid(Now(), 11, 6)) < "10:00" Then
    Beep
    MsgBox "Sei in ritardo! Dovresti ricuperare il tempo stasera."
End If

Remind

Err_DLL_Not_Registered:
         ' Check to see if error 429 occurs
         If Err.Number = 429 Then
            MsgBox "Attempting To Register MyServerObject"

            'RegMyDLLAttempted is used to determine whether an
            'attempt to register the ActiveX DLL has already been
            'attempted. This helps to avoid getting stuck in a loop if
            'the ActiveX DLL cannot be registered for some reason.
            If RegMyDLLAttempted Then
               MsgBox "Unable to Register MyServerObject"
              Resume Next
            Else
              RegMyServerObject   'Declared in Module1
               RegMyDLLAttempted = True
               MsgBox "Registration of MyServerObject attempted."
               Resume
            End If
         Else

MsgBox "An Error Occurred"
         End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
         Single, Y As Single)
      'this procedure receives the callbacks from the System Tray icon.
      Dim Result As Long
      Dim msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        msg = X
       Else
        msg = X / Screen.TwipsPerPixelX
       End If
       Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mPopupSys
       End Select
      End Sub
   
      Private Sub Form_Resize()
       'this is necessary to assure that the minimized window is hidden
       If Me.WindowState = vbMinimized Then Me.Hide
      End Sub
   
      Private Sub Form_Unload(Cancel As Integer)
       'this removes the icon from the system tray
       Shell_NotifyIcon NIM_DELETE, nid
      End Sub
   



Private Sub mnuGeoSol_Click()
Dim URL As String
Dim IE As Object

Set IE = CreateObject("InternetExplorer.Application")

URL = "http://www.geosole.info"
    
    IE.Visible = True
    IE.Navigate (URL)
End Sub

      Private Sub mPopExit_Click()
       'called when user clicks the popup menu Exit command
       Unload Me
      End Sub
   
      Private Sub mPopRestore_Click()
       'called when the user clicks the popup menu Restore command
       Me.WindowState = vbNormal
       Result = SetForegroundWindow(Me.hwnd)
       Me.Show
      End Sub



