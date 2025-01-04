' Array of responses from Sunny
Dim responses
responses = Array("Hello!", "How can I help you?", "Nice to meet you!", "What's your name?", "I'm Sunny!", "Do you like birds?", "Tell me more!", "That's interesting!", "Have a great day!", "Goodbye!", "What's your favorite color?", "Do you have any pets?", "I love chatting with you!", "Tell me a fun fact!", "What's your favorite hobby?", "Do you want to play games?", "What's your favorite food?", "Have you traveled anywhere recently?", "What's your favorite movie?", "Do you enjoy reading?", "What's your favorite book?", "Do you like music?", "What's your favorite song?", "Do you play any instruments?", "What's your favorite sport?", "Do you have any siblings?", "What's your dream job?", "Do you like to cook?", "What's your favorite season?", "Do you have any pets?", "What's your favorite animal?")

' Function to display the bird image and add buttons
Sub ShowBirdImage()
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.Navigate "about:blank"
    objIE.Document.Title = "Sunnybird.exe"
    objIE.document.write "<html><head><title>Sunnybird.exe</title></head><body><h1>App Version: 1.0</h1><p>The open chat and share image button don't work</p><img src='https://petco.scene7.com/is/image/PETCO/112151-center-1?$PLP-category$' alt='Sunny'><br><button onclick='window.external.OpenChatWindow()'>Open Chat</button><br><button onclick='window.external.OpenDownloadsFolder()'>Share Image</button></body></html>"
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
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run userMessage
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
