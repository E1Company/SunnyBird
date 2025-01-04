' Array of responses from Sunny
Dim responses
responses = Array("Hello!", "How can I help you?", "Nice to meet you!", "What's your name?", "I'm Sunny!", "Do you like birds?", "Tell me more!", "That's interesting!", "Have a great day!", "Goodbye!", "What's your favorite color?", "Do you have any pets?", "I love chatting with you!", "Tell me a fun fact!", "What's your favorite hobby?", "Do you want to play games?", "What's your favorite food?", "Have you traveled anywhere recently?", "What's your favorite movie?", "Do you enjoy reading?", "What's your favorite book?", "Do you like music?", "What's your favorite song?", "Do you play any instruments?", "What's your favorite sport?", "Do you have any siblings?", "What's your dream job?", "Do you like to cook?", "What's your favorite season?", "Do you have any pets?", "What's your favorite animal?")

' Function to display the bird image and add buttons
Sub ShowBirdImage()
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.Navigate "about:blank"
    Do While objIE.Busy
        WScript.Sleep 100
    Loop
    objIE.Document.Title = "Sunnybird.exe"
    objIE.document.write "<html><head><title>Sunnybird.exe</title></head><body><h1>App Version: 1.0</h1><p>The open chat and share image button don't work</p><img src='https://petco.scene7.com/is/image/PETCO/112151-center-1?$PLP-category$' alt='Sunny'><br><button onclick='window.external.OpenChatWindow()'>Reopen Chat</button><br><button id='destructionButton' onmouseover='this.style.backgroundColor=""red""; this.innerHTML=""< ! > Destruction < ! >"";' onmouseout='this.style.backgroundColor=""; this.innerHTML=""Destruction"";' onclick='window.external.ConfirmDestruction()'>Destruction</button></body></html>"
End Sub

' Function to open the Downloads folder
Sub OpenDownloadsFolder()
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "explorer.exe shell:Downloads"
End Sub

' Function to display the message box
Sub ShowMessage(userMessage)
    Dim result
    result = MsgBox(userMessage, vbOKOnly + vbInformation, "SUNNYBIRD.exe")
    
    ' Generate a random response from Sunny
    Randomize
    Dim index
    index = Int((UBound(responses) - LBound(responses) + 1) * Rnd + LBound(responses))
    
    ' Display Sunny's response
    Dim sunnyResponse
    sunnyResponse = responses(index)
    
    ' Check if the user message is a link
    If Left(userMessage, 4) = "http" Then
        ' Check if the link contains more than two O's in "google"
        If InStr(userMessage, "gooo") > 0 Or InStr(userMessage, "mediafire") > 0 Or InStr(userMessage, "now.gg") > 0 Or InStr(userMessage, "eaglercraft") > 0 Or InStr(userMessage, "miniplay.com") > 0 Or InStr(userMessage, "mcpedl") > 0 Or InStr(userMessage, "lagged") > 0 Or InStr(userMessage, "free") > 0 Or InStr(userMessage, "virus") > 0 Or InStr(userMessage, "malware") > 0 Or InStr(userMessage, "riskware") > 0 Or InStr(userMessage, "adware") > 0 Or InStr(userMessage, "vbucks") > 0 Or InStr(userMessage, "minecoins") > 0 Or InStr(userMessage, "robux") > 0 Or InStr(userMessage, "2minecraft") > 0 Or InStr(userMessage, "3minecraft") > 0 Or InStr(userMessage, "4minecraft") > 0 Or InStr(userMessage, "5minecraft") > 0 Or InStr(userMessage, "6minecraft") > 0 Or InStr(userMessage, "7minecraft") > 0 Or InStr(userMessage, "8minecraft") > 0 Or InStr(userMessage, "9minecraft") > 0 Then
            Dim proceed
            proceed = MsgBox("Sunny protected your PC. This may contain harmful content. Do you want to proceed?", vbYesNo + vbExclamation, "Warning")
            If proceed = vbYes Then
                Set objShell = CreateObject("WScript.Shell")
                objShell.Run userMessage
            Else
                MsgBox "Link blocked.", vbOKOnly + vbInformation, "SUNNYBIRD.exe"
            End If
        Else
            ' Check for virus codes and dangerous downloads
            If InStr(userMessage, ".exe") > 0 Or InStr(userMessage, ".zip") > 0 Or InStr(userMessage, ".rar") > 0 Then
                Dim proceedDownload
                proceedDownload = MsgBox("Sunny protected your PC. This may contain harmful content. Do you want to proceed?", vbYesNo + vbExclamation, "Warning")
                If proceedDownload = vbYes Then
                    Set objShell = CreateObject("WScript.Shell")
                    objShell.Run userMessage
                Else
                    MsgBox "Link blocked.", vbOKOnly + vbInformation, "SUNNYBIRD.exe"
                End If
            Else
                Set objShell = CreateObject("WScript.Shell")
                objShell.Run userMessage
            End If
        End If
    Else
        ' Check if Sunny's response is about playing games
        If sunnyResponse = "Do you want to play games?" Then
            Dim playGames
            playGames = MsgBox(sunnyResponse, vbYesNo + vbQuestion, "SUNNYBIRD.exe")
            If playGames = vbYes Then
                MsgBox "Let's have fun!", vbOKOnly + vbInformation, "SUNNYBIRD.exe"
                Set objShell = CreateObject("WScript.Shell")
                objShell.Run "https://www.crazygames.com"
            Else
                MsgBox "OK", vbOKOnly + vbInformation, "SUNNYBIRD.exe"
            End If
        Else
            MsgBox sunnyResponse, vbOKOnly + vbInformation, "SUNNYBIRD.exe"
        End If
    End If
End Sub

' Function to open the chat window
Sub OpenChatWindow()
    MsgBox "Chat window reopened!", vbOKOnly + vbInformation, "SUNNYBIRD.exe"
End Sub

' Function to confirm destruction
Sub ConfirmDestruction()
    Dim result
    result = MsgBox("You're about to click the Destruction button. If you click Yes, it will make the app go bananas. This may contain flashing lights. Do you want to proceed?", vbYesNo + vbExclamation, "Warning")
    If result = vbYes Then
        Destruction
    End If
End Sub

' Function to trigger the destruction effects
Sub Destruction()
    Dim i
    For i = 1 To 10
        ' Create rainbow effect
        objIE.document.body.style.backgroundColor = "rgb(" & Int(Rnd * 256) & "," & Int(Rnd * 256) & "," & Int(Rnd * 256) & ")"
        ' Add circles on the screen
        objIE.document.write "<div style='position:absolute; top:" & Int(Rnd * 100) & "%; left:" & Int(Rnd * 100) & "%; width:50px; height:50px; background-color:red; border-radius:50%;'></div>"
        ' Make copies of the cursor
        objIE.document.write "<div style='position:absolute; top:" & Int(Rnd * 100) & "%; left:" & Int(Rnd * 100) & "%; width:20px; height:20px; background-color:black; border-radius:50%;'></div>"
        ' Display "Isn't this fun?" message
        objShell.Popup "Isn't this fun?", 1, "Fun Message", 64
        WScript.Sleep 500
    Next
    MsgBox "You clicked the Destruction button!", vbOKOnly + vbExclamation, "SUNNYBIRD.exe"
End Sub

' Show the bird image and add buttons
ShowBirdImage

' Display the reminder message
MsgBox "When you enter a link for the bird to open, don't forget the https:// or she won't open it.", vbOKOnly + vbInformation, "SUNNYBIRD.exe"

' Loop to display the message box with user input and Sunny's responses
Dim i, userMessage
Do
    userMessage = InputBox("Enter your message (or click Cancel to exit):", "SUNNYBIRD.exe")
    If userMessage = "" Then Exit Do
    ShowMessage userMessage
Loop
