Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Form      : modTimedMsgBox
' DateTime  : 14/03/2026
' Author    : Colin Riddington (Mendip Data Systems)
' Website   : https://www.isladogs.co.uk
' Purpose   : Functions used to create a timeout version of the new style Access message box function
'             Also works for old style message box wityh bold first line and standard message box
' Copyright : The code in the utility MAY be altered and reused in your own applications
'             provided the copyright notice is left unchanged (including Author, Website and Copyright)
'             You are NOT allowed to sell, resell or repost this on other sites such as online forums
'             without permission from the author. However, links back to the above website ARE allowed.

'             If you find this code useful please place a link to my website on your own web site
'             so that others may benefit as well.
' Updated   : March 2026
'---------------------------------------------------------------------------------------

'===========================
'  Win32 API Declarations
'===========================
'The following APIs are used to identify the handle for the message box and set the timer
'Finally the timer is destroyed if the message box is closed by the user

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindoww
'Retrieves a handle to the top-level window whose class name and window name match the specified strings.
'Also works for unicode strings
Private Declare PtrSafe Function FindWindowW Lib "user32" _
    (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagew
'Places (posts) a message in the message queue associated with the thread that created the specified window
'and returns without waiting for the thread to process the message.
Private Declare PtrSafe Function PostMessageW Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal wMsg As Long, _
     ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-settimer
'Creates a timer with the specified time-out value.
Private Declare PtrSafe Function SetTimer Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
     ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-killtimer
'Destroys the specified timer
Private Declare PtrSafe Function KillTimer Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long


'https://learn.microsoft.com/en-us/windows/win32/winmsg/wm-close
'Sent as a signal that a window or an application should terminate.
Private Const WM_CLOSE As Long = &H10

'---------------------------------------------------------------------------------------

'===========================
' Module-level state variables
'===========================
Private mTimedMsgTitle As String
Private mTimedResult As VbMsgBoxResult
Private mTimerFired As Boolean

'===========================
' Timed MsgBox function (Eval-wrapped)
'===========================

Public Function MsgBoxT( _
        ByVal Prompt As String, _
        Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
        Optional ByVal Title As String = "", _
        Optional ByVal HelpFile As String = "", _
        Optional ByVal Context As Long = 0, _
        Optional ByVal Timeout As Long = 0) As VbMsgBoxResult
        
    'Reset module level state values (needed for chained messages)
    mTimerFired = False
    mTimedResult = GetDefaultButtonResult(Buttons)
    mTimedMsgTitle = vbNullString
    ' *********************

    Dim EvalString As String
    Dim EscPrompt As String
    Dim EscTitle As String
    Dim EscHelp As String
    Dim r As VbMsgBoxResult
    Dim tID As LongPtr
    
    'if title missing, set to AppTitle or application name
    If Title = "" Then Title = GetAppTitle

    ' Escape quotes for Eval
    EscPrompt = Replace(Prompt, """", """""")
    EscTitle = Replace(Title, """", """""")
    EscHelp = Replace(HelpFile, """", """""")
    
    ' Prepare state
    mTimedMsgTitle = Title
    
    ' Start timer if timeout > 0
    If Timeout > 0 Then
            tID = SetTimer(0, 0, Timeout, AddressOf CloseMsgBoxTimer)
    End If
    
    ' Build Eval string
    If Context < 0 Then 'use bold first line style message
       EvalString = "MsgBox(""" & EscPrompt & """, " & CLng(Buttons) & "," & _
            " """ & EscTitle & """ )"
    Else 'use new style message
        EvalString = "MsgBox(""" & EscPrompt & """, " & CLng(Buttons) & "," & _
            " """ & EscTitle & """, """ & EscHelp & """," & CLng(Context) & ")"
    End If
    
    ' Show new-style Access message box
    r = Eval(EvalString)
    
    ' Kill timer (so it doesn't persist with chained messages)
    KillTimer 0, tID

    ' If timeout fired, return default button
    If mTimerFired Then
        TempVars!Timeout = "Yes"
        MsgBoxT = mTimedResult
    Else
        TempVars!Timeout = "No"
        MsgBoxT = r
    End If
    
    'display returned value (optional)
   ' Debug.Print MsgBoxT
End Function

'===========================
'  Timer callback
'===========================
Public Sub CloseMsgBoxTimer(ByVal hwnd As LongPtr, _
                            ByVal uMsg As Long, _
                            ByVal idEvent As LongPtr, _
                            ByVal dwTime As Long)

    Dim hMsg As LongPtr
    
    'get handle of the message box from its title using FindWindowW
    'this also works with Unicode strings
    'must 'convert' both arguments using StrPtr
    hMsg = FindWindowW(StrPtr(vbNullString), StrPtr(mTimedMsgTitle))
    
    'alternative equivalent code
   ' hMsg = FindWindowW(0, StrPtr(mTimedMsgTitle))
   
   ' Debug.Print hMsg
    
    If hMsg <> 0 Then
        mTimerFired = True
        PostMessageW hMsg, WM_CLOSE, 0, 0
    End If
  
    KillTimer 0, idEvent
End Sub

'===========================
'  Determine default button result
'===========================

Function GetDefaultButtonResult(Buttons As VbMsgBoxStyle) As VbMsgBoxResult
    Dim defBtn As Long
    defBtn = Buttons And &H300 ' mask default button bits
        
    Select Case Buttons And &HF ' button group
        Case vbOKOnly
            GetDefaultButtonResult = vbOK

        Case vbOKCancel
            If defBtn = vbDefaultButton2 Then
                GetDefaultButtonResult = vbCancel
            Else
                GetDefaultButtonResult = vbOK
            End If

        Case vbYesNo
            If defBtn = vbDefaultButton2 Then
                GetDefaultButtonResult = vbNo
            Else
                GetDefaultButtonResult = vbYes
            End If

        Case vbYesNoCancel
            Select Case defBtn
                Case vbDefaultButton2: GetDefaultButtonResult = vbNo
                Case vbDefaultButton3: GetDefaultButtonResult = vbCancel
                Case Else:             GetDefaultButtonResult = vbYes
            End Select

        Case vbRetryCancel
            If defBtn = vbDefaultButton2 Then
                GetDefaultButtonResult = vbCancel
            Else
                GetDefaultButtonResult = vbRetry
            End If

        Case vbAbortRetryIgnore
            Select Case defBtn
                Case vbDefaultButton2: GetDefaultButtonResult = vbRetry
                Case vbDefaultButton3: GetDefaultButtonResult = vbIgnore
                Case Else:             GetDefaultButtonResult = vbAbort
            End Select

        Case Else
            GetDefaultButtonResult = vbOK
    End Select
End Function

'===========================
' Helper Function to set default title if blank
'===========================

Public Function GetAppTitle() As String
    Dim Db As DAO.Database, prp As Property
    
    On Error GoTo Err_Handler
   
    Set Db = CurrentDb
    GetAppTitle = Db.Properties("AppTitle")

Exit_Handler:
   Exit Function

Err_Handler:
    Select Case Err.Number
    Case 3270 'Property Not Found
        'db doesn't have an app title
        GetAppTitle = "Microsoft Access"
    Case Else
        VBA.MsgBox "Error " & Err.Number & " " & Err.Description & " in procedure GetAppTitle", vbCritical, "GetAppTitle error"
        
    End Select
    
    Resume Exit_Handler

End Function
