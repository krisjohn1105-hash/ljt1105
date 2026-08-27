' ==============================================================================
' Windows 작업 스케줄러 백그라운드 무음(Silent) 실행 스크립트
' CMD/콘솔 창이 화면에 전혀 나타나지 않고 완전 백그라운드로 실행되도록 보장합니다.
' ==============================================================================
Dim WshShell, fso, scriptDir, pyScript, pythonwPath, userProfile

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pyScript = scriptDir & "\auto_git_sync.py"

userProfile = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
pythonwPath = userProfile & "\AppData\Local\Programs\Python\Python313\pythonw.exe"

If Not fso.FileExists(pythonwPath) Then
    pythonwPath = "pythonw.exe"
End If

' 0 = 창 숨김 (vbHide), False = 비동기 실행
WshShell.Run """" & pythonwPath & """ """ & pyScript & """ --once", 0, False
