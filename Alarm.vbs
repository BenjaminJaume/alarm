Dim minutes
minutes = 30 'minutes

Function firstDialogBox()
    answer = msgbox("Do you want to start the alarm?" & vbnewline & vbnewline & _
    "Default interval: " & minutes & " minutes" & vbnewline & vbnewline & _
    "YES to start the alarm" & vbnewline & _
    "NO to change the interval" & vbnewline & _
    "CANCEL to stop the alarm", 3, "Starting the alarm")

    if answer = 6 then
        alarm()
    end if

    if answer = 7 then
        errorCode = changeInterval()
        if errorCode = "canceled" then 
            firstDialogBox()
        else 
            alarm()
        end if
    end if
End Function

Function alarm()
    do
        if errorCode <> "canceled" then
            msgbox " Starting alarm for " & minutes & " minutes", 64, "Alarm"
            WScript.Sleep minutes * 60 * 1000
        end if

        answerTimesUpDialogBox = MsgBox("Time for a mindful break" _
            & vbnewline & vbnewline & _
            "Current interval: " & minutes & " minute(s)" & vbnewline & vbnewline & _
            "- Be present here and now" & vbnewline & _
            "- Stay off the screen for a bit" & vbnewline & _
            "- Stretch your neck" & vbnewline & _
            "- Gaze around see what's going on" & vbnewline & _
            "- Need to go to the toilet?" & vbnewline & _
            "- Get up and jump, dance, smile!" & vbnewline & _
            "- Get some water" & vbnewline & _
            "- Stretch that out" & vbnewline & _
            "- Go for a small walk" & vbnewline & vbnewline & _
            "YES to restart the alarm" & vbnewline & _
            "NO to change the interval" & vbnewline & _
            "CANCEL to stop the alarm", 3, "Alarm - Time's up!")

        if answerTimesUpDialogBox = 6 then
            errorCode = ""
        end if
        if answerTimesUpDialogBox = 7 then
            errorCode = changeInterval()
            if errorCode = "canceled" then

            end if
        else if answerTimesUpDialogBox = 2 then
                msgbox " The alarm has been stopped", 64, "Stoping alarm"
                exit do 'exit loop
            end if
        end if
    loop
End Function

Function changeInterval()
    newMinutes = InputBox("Enter alarm interval in minutes" & vbnewline & vbnewline & "Current interval: " & minutes & " minutes", "Set new alarm time") 'minutes
    
    do until IsNumeric(newMinutes) = true
        newMinutes = InputBox("Invalid Entry. Enter alarm interval in minutes" & vbnewline & vbnewline & "Current interval: " & minutes & " minutes", "Set new alarm time") 'minutes
    loop

    if IsEmpty(newMinutes) then
        changeInterval = "canceled" 'returning 1 if user hits CANCEL
    else
        minutes = newMinutes
        ' msgbox " New interval set for " & minutes & " minutes", 64, "New interval"
    end if
End Function

firstDialogBox()