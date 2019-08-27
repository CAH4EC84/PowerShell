function getMail { #Обрабочик прайсов с почты

    $sqlCmd.CommandText = "select id,node,fromPath,toPath,toDir,isArc, vcs from [IncomingPrices] where enabled = 1  and  type = 'mail' order by id desc"
    $sqlAdapter =  New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
    $sqlAdapter.Fill($dataSet) | Out-Null  
    "MAIL Start at : {0}" -f (Get-Date).Tostring()  | Out-File -FilePath $logPath -Append    
    While ($inboxFolder.Items.count -gt 0) { # перебираем входящие письма
    "MAIL check items : {0}" -f $inboxFolder.Items.count  | Out-File -FilePath $logPath -Append

        foreach($item in $inboxFolder.items) {
        "MAIL check item.TO : {0}" -f $item.to  | Out-File -FilePath $logPath -Append
    
            If ($item.attachments.count -gt 0) {            
                foreach($recipient in $item.recipients){
                #Проверяем регистрацию адреса получаетля и соответсвие архив \ или просто файл.
                    $validRow=""
                    $checkMail ="fromPath='" + $recipient.AddressEntry.Address + "'"
                    $validRow = $dataSet.Tables[0].Select($checkMail)
                    if ($validRow.count -eq 1) {
                        "MAIL valid recipient found  id: {0} checkMail: {1}" -f $validRow.id,$checkMail  | Out-File -FilePath $logPath -Append
                        foreach ($attach in $item.Attachments) {
                            $attachExtension=""
                            $attachExtension = (Split-Path -Path $attach.FileName -Leaf).Split('.')[-1]  
                            if ( 
                                 ( ($validRow[0].isArc -eq 0) -and ($attachExtension -in ($validFormats)) ) -or ( ($validRow[0].isArc -eq 1) -and ($attachExtension -in ($validArcFormats)) )
                               ) {
                                     #Если нет  целевой и архивной папки то создаем
                                    if (!(Test-Path -path ($validRow.toDir + "\" + $cDate + "\"))) {
                                        New-Item ($validRow.toDir + "\" + $cDate  + "\") -Type Directory
                                    }
                                    "MAIL valid attachment found  id: {0} attach: {1}" -f $validRow.id,$attachExtension  | Out-File -FilePath $logPath -Append
                                    savePrice -from "MAIL" -refreshed 1
                            }
                        } 
                    }
                } 
                
                $item.move($arcMailFolder)
            } else {
                "MAIL no attachments item.TO {0}" -f $item.to  | Out-File -FilePath $logPath -Append
                ($item).delete()
            }
        }
    }
}

function getFtp {#Обрабочик прайсов с фтп
    $sqlCmd.CommandText = "select id,node,fromPath,toPath,toDir,isArc, vcs from [IncomingPrices] where enabled = 1 and type='FTP' order by id desc"
    $sqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
    $sqlAdapter.Fill($dataSet) | Out-Null
    "FTP Script start at :{0}" -f (Get-Date).Tostring()  | Out-File -FilePath $logPath -Append
    foreach ($row in $dataSet.Tables[0].Rows) {
        
        #Проверям существование исходного файла        
        IF (Test-Path $row.fromPath ){
          $from = Get-Item($row.fromPath)
          "FTP Check: {0} : {1} File found " -f $row.id, $row.fromPath | Out-File -FilePath $logPath -Append
        } 
        ELSE {#Регистрируем в логе делаем пометку и переходим к следующей строке
           "FTP Check: {0} : {1} File not found " -f $row.id, $row.fromPath | Out-File -FilePath $logPath -Append

           continue 
        }

        #Если нет  целевой и архивной папки то создаем
        if (!(Test-Path -path ($row.toDir + "\" + $cDate + "\"))) {
            New-Item ($row.toDir + "\" + $cDate  + "\") -Type Directory
        }

        #Проверяем есть более старый файл в целевой папке и перемещаем его в архивную
        

        #Сохраняем \ распаковываем \ переименовываем новый прайс в целевой папке 
        savePrice -from "FTP" -refreshed (rotateArc)

    }
}




function rotateArc {
 
    $arcFile=""
    $arcFile=$row.toDir + '\Archive' + (Get-Item($row.fromPath)).Extension  
    IF (Test-Path $arcFile ) {
    
      #если на фтп есть более новый
      IF ( (Get-Item $arcFile).LastWriteTime.Ticks -lt (Get-Item $row.fromPath).LastWriteTime.Ticks ) {
        "FTP Check: {0} : {1} External file refreshed " -f $row.id, $row.fromPath | Out-File -FilePath $logPath -Append
        Copy-Item -Path $arcFile -Destination (`
                                                $row.toDir +'\' + $cDate +'\' `
                                                +(Get-Item $arcFile).LastWriteTime.Ticks `
                                                +"_" + (Get-Item $row.fromPath).Name )  
        return 1
      } Else {
        "FTP Check: {0} : {1} External file has same date " -f $row.id, $row.fromPath | Out-File -FilePath $logPath -Append
        return 0 
      }

    } Else { #Если архивного нет то копируем
        
        "FTP Check: {0} : {1} No Internal copy External file" -f $row.id, $row.fromPath | Out-File -FilePath $logPath -Append
        return 1
    }


}

function savePrice ([string] $from,[int] $refreshed) {
    IF ($from -eq "FTP") {
   #Сохраняем с фтп
        if ($refreshed -eq 1)    {
            
            Copy-Item -Path $row.fromPath -Destination ($row.toDir + '\Archive' + (Get-Item($row.fromPath)).Extension) -Force  
            if ($row.isArc -eq 1) {
                expandArchive -Path ($row.toDir + '\Archive' + (Get-Item($row.fromPath)).Extension)  -Destination $row.toDir
                Move-Item -Path ($row.toDir+'\*.' +(Split-Path -Path $row.toPath -Leaf).Split('.')[-1]) $row.toPath -Force
            } else {
                Copy-Item -Path  ($row.toDir + '\Archive' + (Get-Item($row.fromPath)).Extension) $row.toPath -Force
            }
            updateTimeOf -rowId $row.id
            "FTP Check: {0} : {1} Internal file refreshed" -f $row.id, $row.toPath | Out-File -FilePath $logPath -Append
        }
    } ELSE {
    #Сохраныем с почты
        #Сохраняем вложение как archive в рабочую и и архивную директорию
        $tmpFile=""
        $tmpFile = ($validRow.toDir+'\Archive.'+$attachExtension)
        $attach.SaveAsFile($tmpFile)
        "MAIL attachment saved  id: {0} attach: {1}" -f $validRow.id,$tmpFile  | Out-File -FilePath $logPath -Append
        Copy-Item -Path $tmpFile -Destination ($validRow.toDir +'\' + $cDate +'\' `
                                                +(Get-Item $tmpFile).LastWriteTime.Ticks `
                                                +"_" + (Get-Item $tmpFile).Name )
        
        if ($validRow.isArc -eq 1) {
            expandArchive -Path $tmpFile  -Destination $validRow.toDir
            Move-Item -Path ($validRow.toDir+'\*.' +(Split-Path -Path $validRow.toPath -Leaf).Split('.')[-1]) $validRow.toPath -Force
            
        } else {
                Copy-Item -Path  $tmpFile $validRow.toPath -Force
            }
            "MAIL arc expanded or file renamed id: {0}" -f $validRow.id | Out-File -FilePath $logPath -Append
            updateTimeOf -rowId $validRow.id
        
        
    
    }
}

function expandArchive([string]$Path, [string]$Destination) {
	
	$7zArgs = @(
        'e'
		'-y'						## assume Yes on all queries
		"`"-o$($Destination)`""		## set Output directory
		"`"$($Path)`""				## <archive_name>
	)
	&$7zApp $7zArgs > null
}

function updateTimeOf ([int] $rowId) {
    $sqlCmd = $SqlConnection.CreateCommand()
    $sqlCmd.CommandText = "Update [IncomingPrices] set lastUpdate = getDate() where id ="+ $rowId
    $rowsAffected = $sqlCmd.ExecuteNonQuery()
}


#Основной модуль обработки вхрдящих прайсов#
TRY {

#Пременные
$LOCAL:7zApp = "c:\Priceimp\7-Zip\7z.exe"

$LOCAL:cDate =  (Get-Date -UFormat "%Y.%m.%d")
$LOCAL:logPath = "C:\PriceImp\#PriceRobot\Prices.log"
#Если нет лога создаем его
if (!(Test-Path -path ($logPath))) {New-Item ($logPath) -Type File} 


#SQL 
$LOCAL:sqlServer = "meddb"
$LOCAL:sqlCatalog = "medisTest"
$LOCAL:sqlLogin = "sa"
$LOCAL:sqlPassw = "supersecretpassword"
$LOCAL:sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$LOCAL:sqlConnection.ConnectionString = "Server=$SqlServer; Database=$SqlCatalog; User ID=$SqlLogin; Password=$SqlPassw;"
$LOCAL:sqlConnection.Open()
$LOCAL:sqlCmd = $SqlConnection.CreateCommand()
$LOCAL:dataSet = New-Object System.Data.DataSet

#Outlook
$LOCAL:outlook = New-Object -ComObject Outlook.Application
$LOCAL:nameSpace = $outlook.GetNameSpace("MAPI")
$LOCAL:accFolder = $nameSpace.Folders("medlineuser@price.medline.spb.ru")
$LOCAL:inboxFolder = $accFolder.folders("Входящие")
$LOCAL:arcMailFolder = $accFolder.folders("Archive")
$LOCAL:validFormats = 'dbf','xls','xlsx'
$LOCAL:validArcFormats = 'zip','rar','7z'


getFtp
"FTP Finished at : {0}" -f (Get-Date).Tostring()  | Out-File -FilePath $logPath -Append


getMail
"MAIL Finished at : {0}" -f (Get-Date).Tostring()  | Out-File -FilePath $logPath -Append

}

Catch [system.exception] { 
    "FTP Some error {0} `t {1}" -f $Error[0],$Error[0].ScriptStackTrace | Out-File -FilePath $logPath -Append }
Finally {
    $dataSet.Clear()
    $SqlConnection.close()
    $outLook=""
    $nameSpace=""
    "Script stop at {0}" -f (Get-Date).Tostring()  | Out-File -FilePath $logPath -Append
}
