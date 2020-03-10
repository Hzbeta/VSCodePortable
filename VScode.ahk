#Include <DownloadFile>
#SingleInstance ignore
SetWorkingDir, %A_ScriptDir%
SetTitleMatchMode,2
VSCodeWinTitle:="Visual Studio Code ahk_exe Code.exe"
IniRead, LastCheckUpdate, VSCode.ini, Default, LastCheckUpdate, 0
IniRead, CheckUpdateInterval, VSCode.ini, Default, CheckUpdateInterval, 7
IniRead, UpdateAtNextLaunch, VSCode.ini, Default, UpdateAtNextLaunch, 0
IniRead, VSCodeBranch, VSCode.ini, Default, VSCodeBranch, "win32-x64"
global VSCodeWinTitle

Menu,TRAY,Tip,VSCodePortable
Menu,TRAY,NoStandard
Menu,TRAY,Add,About,MenuAbout
Menu,TRAY,Add,Exit,MenuExit

if(A_Language=0804||A_Language=1004){
    ; lang=zh
    s_CheckUpdateFailed:="检查更新失败，将在下次启动时重试"
    s_UpdateAliveable:="发现更新！"
    s_LatestVersion:="最新版本："
    s_NowVersion:="当前版本："
    s_UpdateOrNot:="是否进行更新？"
    s_VSCodeRunning:="VSCode正在运行，无法更新`n按“是”关闭VSCode，或按“否”推迟更新到下次启动时。"
    s_CannotGetLatestVScodeInfo:="无法获得最新版本信息！"
    s_ApplyingUpdate:="正在应用更新..."
    s_CannotDelOldFile:="无法删除旧的VSCode文件夹，请重试！"
    s_UpdateFileNotExistOrVSCodeRunning:="更新文件不存在或VSCode正在运行，请重试！"
    s_UpdateCompleted:="更新完成！"
} else {
    ; lang=en
    s_CheckUpdateFailed:="Check update failed! Program will retry in next launch."
    s_UpdateAliveable:="Update Aliveable! "
    s_LatestVersion:="Latest version: "
    s_NowVersion:="Now version: "
    s_UpdateOrNot:="Update now? "
    s_VSCodeRunning:="VSCode is Running`nPress Yes close it now, or press No to put off the update to next launch."
    s_CannotGetLatestVScodeInfo:="Can not get infomation of latest version!"
    s_ApplyingUpdate:="Applying update..."
    s_CannotDelOldFile:="Can not delete old VSCode folder, please try again."
    s_UpdateFileNotExistOrVSCodeRunning:="Update file not exist or VSCode is running, please try again."
    s_UpdateCompleted:="Update completed!"
}

argvs:=""                ;处理参数
loop, %0% {
    argv:=%A_Index%
    if(RegExMatch(argv, "\s")) {
        argv="%argv%"
    }
    argvs.=argv . A_Space
}

if(UpdateAtNextLaunch)
{
    if(!WinExist(VSCodeWinTitle)){
        DoUpdate()
    }
}

RunVSCode(argvs)

DayGap=%A_YYYY%%A_MM%%A_DD%
EnvSub, DayGap, %LastCheckUpdate%, days
if(LastCheckUpdate==0||DayGap>CheckUpdateInterval) {
    try{
        CheckUpdate()
        IniWrite, %A_YYYY%%A_MM%%A_DD%, VSCode.ini, Default, LastCheckUpdate
    } catch {
        TrayTip, %s_CheckUpdateFailed%, 5, 1
    }
}

ExitApp

RunVSCode(argvs)
{
    IniRead, PortableEnvs, VSCode.ini, Default, PortableEnvs, %A_Space%
    if(PortableEnvs) {
        loop, Parse, PortableEnvs, |
        {
            StringSplit, EnvInfo, A_LoopField, `:
                EnvName:=EnvInfo1
                EnvValue:=EnvInfo2
                SystemEnvValue:=%EnvName%
                NewProcessEnv:=""
                SystemEnvs:=Array()
                if(SystemEnvValue){
                    loop, Parse, SystemEnvValue, `;
                    {
                        SystemEnvs.Insert(A_LoopField, "1")
                        }
                }
                loop, Parse, EnvValue, `<
                {
                    FullPath:=A_ScriptDir . "\Portable\" . A_LoopField
                    if(!SystemEnvValue||!SystemEnvs.HasKey(FullPath)){
                        NewProcessEnv:=FullPath . ";" . NewProcessEnv
                    }
                }
                if(!SystemEnvValue){
                    StringTrimRight, NewProcessEnv, NewProcessEnv, 1
                }
                EnvSet, %EnvName%, %NewProcessEnv%%SystemEnvValue%
                SystemEnvs:=""
            }
        }
        if(argvs){
            Run, VSCode\Code.exe %argvs%, VSCode
        } else {
            Run, VSCode\Code.exe, VSCode
        }
    }
    
    
    
    
    
    CheckUpdate()
    {
        if(IsOnline()){
            Req:=RequestNoRedirects("https://update.code.visualstudio.com/latest/" . VSCodeBranch . "-archive/stable")
            ZipLocation:=Req.getResponseHeader("Location")
            if(Req.status==302&&InStr(ZipLocation, ".zip")){
                RegExMatch(ZipLocation,"^.*-([\d\.]*)\.zip$",LocationMatch)
                LatestVersion:=LocationMatch1
                FileGetVersion, NowVersion, VSCode\Code.exe
                if(VerSub(LatestVersion,NowVersion)>0){
                    MsgBox, 4, %UpdateAliveable%%s_LatestVersion%%LatestVersion%`n%s_NowVersion%%NowVersion%`n%s_UpdateOrNot%
                    IfMsgBox, Yes
                    {
                        DownloadFile(ZipLocation,"update.zip")
                        While(WinExist(VSCodeWinTitle)){
                            MsgBox, 4, Infomation, %s_VSCodeRunning%
                            IfMsgBox, Yes
                            {
                                RunWait,taskkill.exe /f /im conhost.exe /im Code.exe,,Hide
                                continue
                            } else {
                                IniWrite, 1, VSCode.ini, Default, UpdateAtNextLaunch
                                Return
                            }
                        }
                        DoUpdate()
                    }
                }
            } else {
                MsgBox, , Error, %s_CannotGetLatestVScodeInfo%
            }
        }
        Return
    }
    
    
    DoUpdate(){
        If(FileExist("update.zip")&&!WinExist(VSCodeWinTitle))
        {
            TrayTip, VSCode, %s_ApplyingUpdate%, 5, 1
            RunWait,taskkill.exe /f /im conhost.exe,,Hide
            if(FileExist("VSCode")){
                FileMoveDir, VSCode\data, data
                FileRemoveDir, VSCode, 1
            }
            if(FileExist("VSCode")){
                MsgBox, , Error, %s_CannotDelOldFile%
                ExitApp
            }
            RunWait,powershell.exe "Expand-Archive -Path update.zip -DestinationPath VSCode"
            FileMoveDir, data, VSCode\data
            TrayTip, VSCode, %s_UpdateCompleted%, 5, 1
            IniWrite, 0, VSCode.ini, Default, UpdateAtNextLaunch
            FileDelete,update.zip
        } else {
            MsgBox, , Error, %s_UpdateFileNotExistOrVSCodeRunning%
        }
        Return
    }
    
    
    RequestNoRedirects(url) {
        WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        WebRequest.Option(6) := False ; No redirects
        WebRequest.Open("GET", url, false)
        WebRequest.Send()
        Return, WebRequest
    }
    
    
    IsOnline(flag=0x40) {
        Return DllCall("Wininet.dll\InternetGetConnectedState", "Str", flag,"Int",0)
    }
    
    VerSub(VerA, VerB)
    {
        StringSplit, VerAs, VerA, `.
        StringSplit, VerBs, VerB, `.
        Loop % VerAs0>VerBs0 ? VerAs0 : VerBs0
        {
            if(VerAs%A_Index% != VerBs%A_Index%)
            {
            Return (0 . VerAs%A_Index%) - (0 . VerBs%A_Index%)
        }
    }
    Return 0
}

MenuAbout:
    Run,https://github.com/Hzbeta/VSCodePortable
Return

MenuExit:
    ExitApp
Return