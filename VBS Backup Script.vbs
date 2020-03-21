set objShell = CreateObject("wscript.shell")
set objFile = CreateOBject("scripting.FileSystemObject")
Set objPlayer = CreateObject("WMPlayer.OCX")

varPopup = 0
strTitle = objFile.GetBaseName(wscript.ScriptName)
strSourcePath = ""
strBackupPath = ""
varInterval = ""
varSkip = 0
varSkip2 = 0
strPath = objFile.GetParentFolderName(WScript.ScriptFullName)
varCorrupt = 0
strSourceInfo = "Enter the path of the file you want to backup"
strBackupInfo = "Enter the path where you want the backup file to be placed"
strIntervalInfo = "Enter how many minutes you want to backup everytime"
strMojibake2 = ""
strHelpText = "lorem ipsum" & vbCrlf & "sample text"

function error(fmessage)
    if varCorrupt = 1 then
        varAnswer = objShell.popup(strMojibake,1,strTitle)
        wscript.quit
    end if

    max=4
    min=1
    randomize
    varRandomRepeat = Int((max-min+1)*Rnd+min)
    varWhile = 0
    varErrorCode = ""

    while varWhile < varRandomRepeat
        max=4
        min=1
        randomize
        varRnd = Int((max-min+1)*Rnd+min)

        if varRnd = 1 then
            varErrorString = 69
        elseif varRnd = 2 then
            varErrorString = 420
        elseif varRnd = 3 then
            varErrorString = 666
        elseif varRnd = 4 then
            varErrorString = 1337
        end if

        varErrorCode = varErrorCode & varErrorString
        varWhile = varWhile+1
    wend

    varError = MsgBox("Error: " & fmessage,0+16,"Error")
end function

function calculate(fequation)
    varChar = 0
    varFunctionSkip = 0
    varCharSection = 0
    varWhile = 0
    varNumberExists = 0

    while varWhile = 0
        varChar = varChar + 1

        if varCharSection = 0 then
            if mid(fequation,varChar,1) = "/" then
                varCharSection = 1
            else
                varFunctionSkip = 1
                varWhile = 1
            end if
        elseif varCharSection = 1 then
            if isNumeric(mid(fequation,varChar,1)) = false then
                if mid(fequation,varChar,1) = "_" then
                    varCharSection = 2
                end if
            else
                varFunctionSkip = 1
                varWhile = 1
            end if
        elseif varCharSection = 2 then
            if isNumeric(mid(fequation,varChar,1)) = true then
                varNumberExists = 1
            elseif mid(fequation,varChar,1) = "_" and varNumberExists = 1 then
                varCharSection = varCharSection + 1
                varNumberExists = 0
            else
                varFunctionSkip = 1
                varWhile = 1
            end if
        elseif varCharSection = 3 then
            if varChar > len(fequation) and varNumberExists = 1 then
                varWhile = 1
                varNumberExists = 0
            elseif isNumeric(mid(fequation,varChar,1)) = true then
                varNumberExists = 1
            else
                varFunctionSkip = 1
                varWhile = 1
            end if
        end if
    wend

    if varFunctionSkip = 0 then
        varSpace1 = InStr(fequation,"_")
        varSpace2 = InStr(varSpace1+1,fequation,"_")
        strOperation = left(fequation,varSpace1-1)
        varValue1 = int(mid(fequation,varSpace1+1,(varSpace2-varSpace1)-1))
        varValue2 = int(mid(fequation,varSpace2+1,len(fequation)-varSpace2))

        if strOperation = "/add" then
            calculate = varValue1+varValue2
        elseif strOperation = "/subtract" then
            calculate = varValue1-varValue2
        elseif strOperation = "/multiply" then
            calculate = varValue1*varValue2
        elseif strOperation = "/divide" then
            calculate = varValue1/varValue2
        elseif strOperation = "/random" then
            varMax=varValue2
            varMin=varValue1
            randomize
            calculate = (int((varMax-varMin+1)*rnd+varMin))
        end if
    else
        error("Calculation failed")
    end if
end function

varGateB = 0

while varGateB = 0
    if objFile.FolderExists(strPath & "\data") then
        varPopup = objShell.popup("WELCOME TO " & Ucase(strTitle) & vbCrlf & vbCrlf & "Made by Ionvop",3,strTitle,0+64)
        objFile.CreateTextFile(strPath & "\data\confirm.vbs").WriteLine("set objShell = CreateObject(""wscript.shell"")" & vbCrlf & "set objFile = CreateOBject(""scripting.FileSystemObject"")" & vbCrlf & vbCrlf & "strPath = objFile.GetParentFolderName(WScript.ScriptFullName)" & vbCrlf & vbCrlf & "objFile.CreateTextFile(strPath & ""\data.txt"",true).WriteLine(""1"")" & vbCrlf & "varAnswer = MsgBox(""Close this window or click OK to stop the operation"",0+64,""Backup Running"")" & vbCrlf & "objFile.CreateTextFile(strPath & ""\data.txt"",true).WriteLine(""0"")")
        objFile.CreateTextFile(strPath & "\data\data.txt",true).WriteLine("0")
        objFile.CreateTextFile(strPath & "\data\notification.ps1",true).WriteLine("[void] [System.Reflection.Assembly]::LoadWithPartialName(""System.Windows.Forms"")" & vbCrlf & "$objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon" & vbCrlf & "$objNotifyIcon.Icon = ""C:\Windows\WinSxS\amd64_microsoft-windows-sechealthui.appxmain_31bf3856ad364e35_10.0.18362.387_none_089f5ee56fb677c6\Device.contrast-white.ico""" & vbCrlf & "$objNotifyIcon.BalloonTipIcon = ""None""" & vbCrlf & "$objNotifyIcon.BalloonTipText = ""Backup success""" & vbCrlf & "$objNotifyIcon.BalloonTipTitle = """ & strTitle & """" & vbCrlf & "$objNotifyIcon.Visible = $True" & vbCrlf & "$objNotifyIcon.ShowBalloonTip(10000)")
        objFile.CreateTextFile(strPath & "\data\notify.bat",true).WriteLine("powershell -executionpolicy bypass -file """ & strPath & "\data\notification.ps1""" & vbCrlf & "PING -n 10 LOCALHOST" & vbCrlf & """" & strPath & "\data""")
        objFile.CreateTextFile(strPath & "\data\run notification.vbs",true).WriteLine("Set objShell = CreateObject(""WScript.Shell"")" & vbCrlf & "objShell.Run chr(34) & """ & strPath & "\data\notify.bat"" & Chr(34), 0")

        if objFile.FileExists(strPath & "\data\dataaaaa.txt") then
            objFile.DeleteFile strPath & "\data\dataaaaa.txt",true
        end if

        varGateB = 1
    else
        varPopup = objShell.popup("LOADING..." & vbCrlf & vbCrlf & "Made by Ionvop",1,strTitle,0+64)
        objFile.CreateFolder(strPath & "\data")
        objFile.CreateTextFile(strPath & "\data\confirm.vbs").WriteLine("set objShell = CreateObject(""wscript.shell"")" & vbCrlf & "set objFile = CreateOBject(""scripting.FileSystemObject"")" & vbCrlf & vbCrlf & "strPath = objFile.GetParentFolderName(WScript.ScriptFullName)" & vbCrlf & vbCrlf & "objFile.CreateTextFile(strPath & ""\data.txt"",true).WriteLine(""1"")" & vbCrlf & "varAnswer = MsgBox(""Close this window or click OK to stop the operation"",0+64,""Backup Running"")" & vbCrlf & "objFile.CreateTextFile(strPath & ""\data.txt"",true).WriteLine(""0"")")
    end if
wend

do
    strData = objFile.OpenTextFile(strPath & "\data\data.txt",1).ReadAll
    varGate6 = 0

    while varGate6 = 0
        varGate5 = 0

        while varGate5 = 0
            varGate3 = 0

            while varGate3 = 0
                if varSkip = 0 then
                    varGate = 0

                    while varGate = 0
                        strSourcePath = InputBox(strSourceInfo,strTitle,strSourcePath)

                        if varCorrupt = 1 then
                            if Lcase(strSourcePath) = "i hate asuna from sao" then
                                varGate = 1
                            else
                                error("f")
                            end if
                        else
                            if objFile.FileExists(strSourcePath) then
                                strType = "file"
                                varGate = 1
                            else
                                if objFile.FolderExists(strSourcePath) then
                                    strType = "folder"
                                    varGate = 1
                                else
                                    if left(strSourcePath,1) = "/" then
                                        if InStr(strSourcePath,"_") = 0 then
                                        else
                                            varBaseCommandLength = InStr(strSourcePath,"_")-1
                                        end if

                                        if strSourcePath = "/help" then
                                            varAnswer = MsgBox(strHelpText,0+64,strTitle)
                                        elseif strSourcePath = "/subscribe" then
                                            objShell.run("https://www.youtube.com/channel/UCXDfWc9wKYat9KmgRRMqaDg")
                                        elseif left(strSourcePath,varBaseCommandLength) = "/add" or left(strSourcePath,varBaseCommandLength) = "/subtract" or left(strSourcePath,varBaseCommandLength) = "/divide" or left(strSourcePath,varBaseCommandLength) = "/multiply" or left(strSourcePath,varBaseCommandLength) = "/random" then
                                            varAnswer = MsgBox(calculate(strSourcePath),0+64,strTitle)
                                        else
                                            varAnswer = error("Command not found")
                                        end if
                                    else
                                        if strSourcePath = "" then
                                            wscript.quit
                                        else
                                            error("File not found")
                                        end if
                                    end if
                                end if
                            end if
                        end if
                    wend
                end if

                varSkip = 0

                if varSkip2 = 0 then
                    varGate2 = 0

                    while varGate2 = 0
                        strBackupPath = InputBox(strBackupInfo,strTitle,strBackupPath)

                        if varCorrupt = 1 then
                            strData = objFile.OpenTextFile(strPath & "\data\data.txt",1).ReadAll
                            strData = left(strData,4)

                            if strData = "666" then
                            else
                                error("f")
                            end if
                        end if

                        if objFile.FolderExists(strBackupPath) then
                            if InStr(1,strBackupPath,strSourcePath,1) = 0 then
                                varGate2 = 1
                                varGate3 = 1
                            else
                                error("Source path must not be within the backup path")
                            end if
                        else
                            if strBackupPath = "" then
                                varGate2 = 1
                            else
                                error("Folder not found")
                            end if
                        end if
                    wend
                else
                    varGate3 = 1
                end if

                varSkip2 = 0
            wend

            varGate4 = 0

            while varGate4 = 0
                varInterval = InputBox(strIntervalInfo & vbCrlf & vbCrlf & strMojibake2 & "Min = 1" & vbCrlf & strMojibake2 & "Max = 60",strTitle,varInterval)

                if varInterval = "" then
                    varGate4 = 1
                    varSkip = 1
                else
                    if isNumeric(varInterval) then
                        if varInterval = 177013 then
                            if varCorrupt = 1 then
                                if objFile.FileExists(strPath & "\data\dataaaaa.txt") then
                                    objFile.DeleteFile strPath & "\data\dataaaaa.txt",true
                                end if
                                
                                do
                                    objPlayer.URL = "C:\Windows\Media\ringout.wav"
                                    objPlayer.controls.play

                                    while objPlayer.PlayState <> 1
                                        objPlayer.controls.pause
                                        wscript.sleep("10")
                                        objPlayer.controls.play
                                        wscript.sleep("1")
                                    wend

                                    if objFile.FileExists(strPath & "\data\dataaaaa.txt") then
                                        strData = objFile.OpenTextFile(strPath & "\data\dataaaaa.txt",1).ReadAll
                                    end if

                                    if strData = "228922" then
                                        objPlayer.URL = "C:\Windows\Media\Windows Proximity Notification.wav"
                                        objPlayer.controls.play

                                        while objPlayer.PlayState <> 1
                                            wscript.sleep("10")
                                        wend

                                        MsgBox("Easter egg complete")
                                        wscript.quit
                                    end if
                                loop
                            end if
                            varWhile = 0
                            strMojibake = ""

                            while (varWhile < 1337)
                                randomize
                                varRnd2 = int((254-128+1)*rnd+128)
                                strMojibake = strMojibake & chr(varRnd2)
                                varWhile = varWhile+1
                            wend

                            wscript.sleep("3000")
                            objPlayer.URL = "C:\Windows\Media\Alarm05.wav"
                            objPlayer.controls.play
                            wscript.sleep("1000")
                            MsgBox(strMojibake)

                            if objPlayer.PlayState = 1 then
                                objPlayer.URL = "C:\Windows\Media\Windows Proximity Connection.wav"
                                objPlayer.controls.play
                            end if

                            while objPlayer.PlayState <> 1
                                objPlayer.controls.pause
                                wscript.sleep("10")
                                objPlayer.controls.play
                                wscript.sleep("10")
                            wend

                            objPlayer.controls.stop
                            strSourcePath = strMojibake
                            strBackupPath = strMojibake
                            varInterval = strMojibake

                            strSourceInfo = ""
                            varWhile = 0

                            while (varWhile < 69)
                                randomize
                                varRnd2 = int((254-128+1)*rnd+128)
                                strSourceInfo = strSourceInfo & chr(varRnd2)
                                varWhile = varWhile+1
                            wend

                            strBackupInfo = ""
                            varWhile = 0

                            while (varWhile < 69)
                                randomize
                                varRnd2 = int((254-128+1)*rnd+128)
                                strBackupInfo = strBackupInfo & chr(varRnd2)
                                varWhile = varWhile+1
                            wend

                            strIntervalInfo = ""
                            varWhile = 0

                            while (varWhile < 69)
                                randomize
                                varRnd2 = int((254-128+1)*rnd+128)
                                strIntervalInfo = strIntervalInfo & chr(varRnd2)
                                varWhile = varWhile+1
                            wend

                            strTitle = ""
                            varWhile = 0

                            while (varWhile < 69)
                                randomize
                                varRnd2 = int((254-128+1)*rnd+128)
                                strTitle = strTitle & chr(varRnd2)
                                varWhile = varWhile+1
                            wend

                            strMojibake2 = ""
                            varWhile = 0

                            while (varWhile < 69)
                                randomize
                                varRnd2 = int((254-128+1)*rnd+128)
                                strMojibake2 = strMojibake2 & chr(varRnd2)
                                varWhile = varWhile+1
                            wend

                            varCorrupt = 1
                            varGate4 = 1
                        else
                            if varInterval < 1 then
                                error("Below minimum value of 1")
                            else
                                if varInterval > 60 then
                                    error("Above minimum value of 60")
                                else
                                    varGate4 = 1
                                    varGate5 = 1
                                end if
                            end if
                        end if
                    else
                        error("Invalid input")
                    end if
                end if
            wend
        wend

        if varInterval = 1 then
            strGrammar = "minute"
        else
            strGrammar = "minutes"
        end if

        if varCorrupt = 1 then
            varAnswer = objShell.popup(strMojibake,1,strTitle)
            wscript.quit
        end if

        varConfirm = MsgBox("Please review the details before proceeding" & vbCrlf & vbCrlf & "File Type: " & strType & vbCrlf & vbCrlf & "Source Path: " & vbCrlf & strSourcePath & vbCrlf & vbCrlf & "Backup Path: " & vbCrlf & strBackupPath & vbCrlf & vbCrlf & "Backup happens every: " & vbCrlf & varInterval & " " & strGrammar,1+48,strTitle)

        if varConfirm = 1 then
            varGate6 = 1
        else
            varSkip = 1
            varSkip2 = 1
        end if
    wend

    varNotify = MsgBox("Would you like to be notified everytime a backup is successful?",4+32,strTitle)

    if strType = "folder" then
        strBackupFolder = strBackupPath & "\" & objFile.GetFileName(strSourcePath)
    end if

    objShell.run("""" & strPath & "\data\confirm.vbs""")
    varFileSuffix = 1
    strFileBaseName = objFile.GetBaseName(strSourcePath)
    strFileExtension = objFile.GetExtensionName(strSourcePath)
    varTime = 0

    varGate7 = 0

    while varGate7 = 0
        wscript.sleep("1000")
        varTime = varTime-1
        strData = objFile.OpenTextFile(strPath & "\data\data.txt",1).ReadAll
        strData = left(strData,1)

        if strData = "0" then
            strAnswer = MsgBox("Operation successfully terminated",0+64,strTitle)
            varGate7 = 1
        end if

        if varTime < 1 then
            if strType = "file" then
                if objFile.FileExists(strSourcePath) then
                    objFile.CopyFile strSourcePath,strBackupPath & "\" & strFileBaseName & " (" & varFileSuffix & ")." & strFileExtension,true
                else
                    varAnswer = MsgBox("The file seems to have been either deleted, moved, or renamed" & vbCrlf & vbCrlf & "To continue the backup, please return the file to its original location or retry the operation and refill the form to its new location",0+48,strTitle)
                end if
            else
                if strType = "folder" then
                    if objFile.FolderExists(strSourcePath) then
                        objFile.CopyFolder strSourcePath,strBackupFolder & " (" & varFileSuffix & ")",true
                    else
                        varAnswer = MsgBox("The folder seems to have been either deleted, moved, or renamed" & vbCrlf & vbCrlf & "To continue the backup, please return the file to its original location or retry the operation and refill the form to its new location",0+48,strTitle)
                    end if
                end if
            end if

            if varNotify = 6 then
                objShell.run("""" & strPath & "\data\run notification.vbs""")
            end if

            varTime = varInterval*10
            varFileSuffix = varFileSuffix+1
        end if
    wend
loop
