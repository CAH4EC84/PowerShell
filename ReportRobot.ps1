TRY{
    #Получаем список файлов для генерации отчета
    $workDir = 'D:\Golikov_A_S\WorkProjects\PowerShell\ReportRobot\'
    $reportGenerator = $workdir + 'ReportGenerator\unireport.exe'
    $generatorArgs = ' '
    $reportConfigs = $workDir + 'ReportSettings\ON\*'



    $logPath = $workdir+ "ReportRobot.log"
    #Если нет лога создаем его
    if (!(Test-Path -path ($logPath))) {New-Item ($logPath) -Type File} 

    
    If (!(Test-Path $reportsConfigs)) {
        foreach ($dilerFolder in Get-item -Path ($reportConfigs))   {
            "{0} : Start generation for Diler : {1}" -f (Get-Date).Tostring(),$dilerFolder  | Out-File -FilePath $logPath -Append
            $doWeek = ($dilerFolder.FullName +'\DayOfWeek\' + ((get-Date).DayOfWeek.value__)+ '\*.ini')
            
            if ( (Test-Path $doWeek) -and (Get-item -Path $doWeek).count -gt 0) {#Если есть папка DayOfWeek она является предпочитетльной.
                "{0} : doWeek : {1}" -f (Get-Date).Tostring(),$doWeek.count  | Out-File -FilePath $logPath -Append
               
               foreach ( $iniFile in (Get-Item -path $doWeek) ) {
                    Copy-Item -Path $iniFile.FullName -Destination ((get-item -Path $reportGenerator).DirectoryName + '\report.ini')
                    Start-Process -FilePath $reportGenerator -ArgumentList $generatorArgs -Wait
               }

            } else { #Если на текущий день нет особы отчетов то генерируем по файлам лежащим в корневой директории поставщика
                $doDaily = Get-Item -Path ($dilerFolder.FullName+'\*.ini')
                "{0} : doDaily : {1}" -f (Get-Date).Tostring(),$doDaily.count  | Out-File -FilePath $logPath -Append
                
                foreach ($iniFile in $doDaily) {
                    Copy-Item -Path $iniFile.FullName -Destination ((get-item -Path $reportGenerator).DirectoryName + '\report.ini')
                    #Start-Process -FilePath $reportGenerator -ArgumentList $generatorArgs -Wait
               }
            }

            "{0} : Stop generation for Diler : {1}" -f (Get-Date).Tostring(),$dilerFolder  | Out-File -FilePath $logPath -Append
        }
    } else {
        "{0} : ERROR - reportFolder not FOUND : {1}" -f (Get-Date).Tostring(),$reportConfigs  | Out-File -FilePath $logPath -Append
    }

}
CATCH [system.exception] {
    "Some error {0}" -f $Error[0]
    "Some error {0}" -f $Error[0].ToString() | Out-File -FilePath $logPath -Append
}
FINALLY {"Script stop {0}" -f (Get-Date).Tostring()}
