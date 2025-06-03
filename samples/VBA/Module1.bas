Attribute VB_Name = "Module1"
Option Explicit

' --- Global Variables for the Timer ---
Public g_TotalSeconds As Long
Public g_TimerRunning As Boolean
Public g_TimeScheduled As Date ' Stores the exact time for the next OnTime event
Public g_PausedTime As Date ' Stores the time when timer was paused
Public g_RemainingOnPause As Long ' Stores remaining seconds when paused
Public g_InitialDurationSeconds As Long ' Stores the last successfully set duration
Public Const DEFAULT_TIMER_DURATION_SECONDS As Long = 3600 ' Default to 1 hour (3600 seconds)

' --- Timer Core Logic ---
Sub ShowTimerForm()
    ' This sub is called from Workbook_Open or frmEmailInput
    Load frmTimer
    frmTimer.Show vbModeless ' Non-modal so other Excel functions (if visible) could be accessed
End Sub

Sub ScheduleNextTick()
    ' Schedules the next call to UpdateTimerDisplay
    If g_TimerRunning = True Then
        g_TimeScheduled = Now + TimeValue("00:00:01") ' Schedule for 1 second from now
        Application.OnTime g_TimeScheduled, "UpdateTimerDisplay", , True ' Schedule for 1 second, allow overwrite
    End If
End Sub

Sub UpdateTimerDisplay()
    ' This sub is called by Application.OnTime every second
    If g_TotalSeconds > 0 And g_TimerRunning = True Then
        g_TotalSeconds = g_TotalSeconds - 1
        frmTimer.lblTime.Caption = FormatTime(g_TotalSeconds) ' Update the display
        ScheduleNextTick ' Schedule the next tick
    ElseIf g_TotalSeconds = 0 And g_TimerRunning = True Then
        ' Timer has reached zero
        g_TimerRunning = False
        frmTimer.lblTime.Caption = "00:00:00"
        MsgBox "Time's up!", vbInformation, "Timer Alert"
        ' Clean up any pending OnTime event
        On Error Resume Next
        Application.OnTime g_TimeScheduled, "UpdateTimerDisplay", , False
        On Error GoTo 0
    Else
        ' Timer was stopped or paused, ensure no pending OnTime event
        On Error Resume Next
        Application.OnTime g_TimeScheduled, "UpdateTimerDisplay", , False
        On Error GoTo 0
    End If
End Sub

' Function to format total seconds into HH:MM:SS string
Function FormatTime(ByVal TotalSeconds As Long) As String
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long
    Dim remainingSecondsAfterHours As Long

    hours = TotalSeconds \ 3600
    remainingSecondsAfterHours = TotalSeconds Mod 3600

    minutes = remainingSecondsAfterHours \ 60
    seconds = remainingSecondsAfterHours Mod 60

    FormatTime = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(seconds, "00")
End Function

' Function to parse user input string into total seconds
Function ParseDurationInput(ByVal durationString As String) As Long
    Dim parts As Variant
    Dim tempSeconds As Long
    Dim numValue As Double
    Dim lowerCaseInput As String
    Dim i As Long

    ParseDurationInput = 0 ' Default to 0 (invalid)

    durationString = Trim(durationString)
    If durationString = "" Then Exit Function

    lowerCaseInput = LCase(durationString)

    ' --- Try HH:MM:SS or MM:SS format first ---
    If InStr(lowerCaseInput, ":") > 0 Then
        parts = Split(lowerCaseInput, ":")
        Dim allNumeric As Boolean
        allNumeric = True
        For i = LBound(parts) To UBound(parts)
            If Not IsNumeric(parts(i)) Then
                allNumeric = False
                Exit For
            End If
        Next i

        If allNumeric Then
            Select Case UBound(parts)
                Case 1 ' MM:SS
                    tempSeconds = CLng(parts(0)) * 60 + CLng(parts(1))
                Case 2 ' HH:MM:SS
                    tempSeconds = CLng(parts(0)) * 3600 + CLng(parts(1)) * 60 + CLng(parts(2))
            End Select
            If tempSeconds > 0 Then
                ParseDurationInput = tempSeconds
                Exit Function
            End If
        End If
    End If

    ' --- Try Xh, Xhr, Xhours, Xm, Xmin, Xminutes format ---
    If InStr(lowerCaseInput, "h") > 0 Then ' Contains 'h' (for hours or half)
        numValue = Val(lowerCaseInput) ' Val will read numbers until non-numeric char
        If numValue > 0 Then
            ParseDurationInput = CLng(numValue * 3600) ' Convert hours to seconds
            Exit Function
        End If
    ElseIf InStr(lowerCaseInput, "m") > 0 Then ' Contains 'm' (for minutes)
        numValue = Val(lowerCaseInput)
        If numValue > 0 Then
            ParseDurationInput = CLng(numValue * 60) ' Convert minutes to seconds
            Exit Function
        End If
    End If

    ' --- If only a number is entered, assume it's in minutes ---
    If IsNumeric(lowerCaseInput) Then
        ParseDurationInput = CLng(lowerCaseInput) * 60
        Exit Function
    End If
End Function

' --- Email Functionality ---

' Function to perform basic email format validation
Function ValidateEmail(ByVal email As String) As Boolean
    ' Basic email validation: check for @ and at least one . after @, and min length
    ValidateEmail = False ' Default to false

    If InStr(email, "@") > 1 Then ' Check for @ not at the beginning
        If InStr(InStr(email, "@") + 1, email, ".") > 0 Then ' Check for . after @
            If Len(email) > 5 Then ' Check for a minimum length
                ValidateEmail = True
            End If
        End If
    End If
End Function

' Subroutine to send the welcome/registration email to the admin
Sub SendWelcomeEmail(ByVal userRegisteredEmail As String)
    Dim olApp As Object ' Represents the Outlook Application
    Dim olMail As Object ' Represents an Outlook MailItem
    Const ADMIN_EMAIL_ADDRESS As String = "ContactCoachVee@gmail.com" ' Your email address

    On Error GoTo ErrorHandler

    ' Try to get a running Outlook instance
    Set olApp = GetObject("Outlook.Application")
    If olApp Is Nothing Then
        ' If no Outlook instance is running, create a new one
        Set olApp = CreateObject("Outlook.Application")
    End If

    ' Create a new email item
    Set olMail = olApp.CreateItem(0) ' 0 corresponds to olMailItem

    With olMail
        .To = ADMIN_EMAIL_ADDRESS
        .Subject = "New Timer App User Registration: " & userRegisteredEmail
        .Body = "A new user has registered for your Excel Timer Application." & vbCrLf & _
                "Registered Email: " & userRegisteredEmail & vbCrLf & _
                "Registration Time (PDT): " & Format(Now, "yyyy-mm-dd hh:mm:ss AM/PM") & vbCrLf & _
                "--------------------------------------------------" & vbCrLf & _
                "This email was sent automatically from the Excel application."
        .Send ' Send the email
    End With

    ' Clean up objects
    Set olMail = Nothing
    Set olApp = Nothing
    Exit Sub ' Exit the sub if successful

ErrorHandler:
    ' Display an error message if something goes wrong (e.g., Outlook not installed/configured)
    MsgBox "Could not send welcome email to admin. Please ensure Outlook is installed and configured " & _
           "on this machine and that security prompts are allowed." & vbCrLf & _
           "Error: " & Err.Description, vbExclamation, "Email Sending Error"

    ' Clean up objects even on error
    If Not olMail Is Nothing Then Set olMail = Nothing
    If Not olApp Is Nothing Then Set olApp = Nothing
End Sub
