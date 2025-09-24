function TriggerMail
{
    param($htmlBody, $username)
    $username = $($env:username)
    $smtpServer = "smtp.na.jnj.com"
    $from = "SA-NCSUS-SQLHELPDESK@its.jnj.com"
    $to = "$($($username -ireplace '^Admin_',''))@its.jnj.com"
    $cc = "PKathirv@its.jnj.com", "DL-ITSUS-GSSQL@ITS.JNJ.COM"
    $subject = "Log Report for the task Start Services : $CRNumber"
    $body = $htmlBody
    $attachment = $Logfile

    Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -CC $cc -Subject $subject -Body $body -Attachments $attachment -BodyAsHtml -UseSsl 
}

function StartServices
{
    param(
        $servername,
        $services_to_start
    )
    try 
    {
        $result = Invoke-Command -ComputerName $servername -ScriptBlock {
            param($services_to_start)
            $sqlservices = Get-Service | Where-Object {$_.DisplayName -like "*SQL*"}
            if($sqlservices)
            {
                $totalservices = $sqlservices.Name
                $before_services_info = New-Object System.Collections.ArrayList
                foreach($service in $totalservices)
                {
                    $before_serviceinfo = Get-Service $service | Select-Object DisplayName, Status, StartType
                    $before_services_info += $before_serviceinfo
                }

            }

            $services = $services_to_start.split(",")
            if($services)
            {
                $timestamp = Get-Date
                $timezone = (Get-TimeZone).id
                $startableservices = @()
                $unstartableservices = @()
                $already_running = @()
                
                $services_info = New-Object System.Collections.ArrayList
                foreach($service in $services)
                {
                    try
                    {
                        $start_service = Get-Service -name $service 
                        if($start_service.Status -eq "Running")
                        {
                            $already_running += $service 
                        }
                        elseif($start_service.Status -eq "Stopped")
                        {
                            Set-Service $service -StartupType Automatic -ErrorAction Stop
                            Start-Service -Name $service -ErrorAction Stop
                            $startableservices += $service
                        }
                        $info = get-service $service
                        $services_info += $info

                    }
                    catch 
                    {
                        $unstartableservices += $service
                    }
                }
                $resultRecord = $services | Where-Object {$_.Status -eq "Stopped"}
                if($resultRecord){$resultRecord = 0}
                else{$resultRecord =  1}
                $resultobject = [PSCustomObject]@{
                    Startservices = $startableservices
                    Unstartservices = $unstartableservices
                    Runningservices = $already_running
                    Resultrecord = $resultRecord
                    TimeStamp = $timestamp
                    TimeZone = $timezone
                    Services_Info = $services_info
                    BeforeServices_Info = $before_services_info
                }
                Return $resultobject
            }
            else
            {
                return $null
            }

        } -ArgumentList $services_to_start

        if($result)
        {
            $connection_failed = @()
            $services_not_started = @()
            $services_started = @()
            Write-Host "Connection to Windows : Established"
            "   Connection to Windows : Established" >> $LogFile
            $fn_htmlBody += "<font size='2'><pre class='tab'>   Connection to Windows     : <font color='green'>Established</font></pre></font>"

            Write-Host "Connection to SQL instance :  Established"
            "   Connection to SQL instance :  Established" >> $LogFile
            $fn_htmlBody += "<font size='2'><pre class='tab'>   Connection to SQL instance: <font color='green'>Established</font></pre></font>"

            Write-Host "ServerTime : $($result.Timestamp) ($($result.TimeZone))"
            "ServerTime : $($result.Timestamp) ($($result.TimeZone))" >> $LogFile
            $fn_htmlBody += "<font size='2'><pre class='tab'>   ServerTime : $($result.Timestamp) ($($result.TimeZone))</pre></font><br>"

            "`n   The status of the services before starting the services" >> $LogFile
            $($result.BeforeServices_Info) | Select-Object DisplayName, Status, StartType >> $LogFile
            
            "`n   The status of the services after starting the services" >> $LogFile
            $($result.Services_Info) | Select-Object DisplayName, Status, StartType >> $LogFile

            #$fn_htmlBody += "<font size='2'><pre class='tab'>   The status of the services are below : </pre></font>"
            
            "   Categories : " >> $LogFile
            $fn_htmlBody += "<font size='2'><pre class='tab'><b>    Categories : <b></pre></font><br>"
           
            Write-Host "    List of services to be started : " -f Cyan
            "       List of services to be started : " >> $LogFile
            $fn_htmlBody += "<font size='2'><pre class='tab1'><b>   List of services to be started : </b></pre></font>"
            $index = 1
            foreach($service in $(($services_to_start).split(",")))
            {
                Write-Host "        $index.$service"
                "           $index.$service" >> $LogFile
                $fn_htmlBody += "<font size='2'><pre class='tab2'>      $index.$service</pre></font>"
                $index++ 
            } 
            $fn_htmlBody += "<br>"
            
            if($result.Startservices) 
            {
                Write-Host "    List of services started : " -f Cyan
                "   List of services started : " >> $LogFile
                $fn_htmlBody += "<font size='2'><br><pre class='tab1'><b>   List of services started : </b></pre></font>"
                $index = 1
                foreach($service in $($result.Startservices))
                {
                    Write-Host "        $index.$service"
                    "           $index.$service" >> $LogFile
                    $fn_htmlBody += "<font size='2'><pre class='tab2'>      $index.$service</pre></font>"
                    $index++ 
                }
                $fn_htmlBody += "<br>"

                        $data = @()
	                    $updateservices = $($result.Startservices) -join ","
	                    $record = [PSCustomObject]@{
		                    ServerName = $servername
		                    Services = $updateservices
	                        }
	                    $data += $record
	                    $data | Export-Csv -Path $csvFilePath -Append -NoTypeInformation
                       
                
            }     
            else
            {
                Write-Host "    List of services started : " -f Cyan -NoNewLine
                Write-Host "NIL" 
                "   List of services started : NIL" >> $LogFile
                $fn_htmlBody += "<font size='2' color='red'><pre class='tab1'><b>   List of services started : NIL</b></pre></font><br>"
                
            }
    
            if($result.Unstartservices)
            {
                Write-Host "    List of services which are not started : " -f Cyan
                "       List of services which are not started : " >> $LogFile
                $fn_htmlBody += "<font size='2' color='red'><pre class='tab1'><b>   List of services which are not started : </b></pre></font><br>"
                $index = 1
                foreach($service in $($result.Unstartservices))
                {
                    Write-Host "        $index.$service"
                    "           $index.$service" >> $LogFile
                    $fn_htmlBody += "<font size='2'><pre class='tab2'>      $index.$service</pre></font>"
                    $index++ 
                }
                $fn_htmlBody += "<br>"

            }
            else
            {
                Write-Host "    List of services which are not started : " -f Cyan -NoNewline
                Write-Host "NIL" 
                "       List of services which are not started : NIL" >> $LogFile
                $fn_htmlBody += "<font size='2'><pre class='tab1'><b>   List of services which are not started : NIL</b></pre></font><br>"
            }
    
            if($result.Runningservices)
            {
                Write-Host "    List of services which are already in running state : " -f Cyan
                "       List of services which are already in running state : " >> $LogFile
                $fn_htmlBody += "<font size='2'><pre class='tab1'><b>   List of services which are already in running state : </b></pre></font><br>"
                $index = 1
                foreach($service in $($result.Runningservices))
                {
                    Write-Host "        $index.$service"
                    "           $index.$service" >> $LogFile
                    $fn_htmlBody += "<font size='2'><pre class='tab2'>      $index.$service</pre></font>"
                    $index++ 
                }
                $fn_htmlBody += "<br>"
            }
            else
            {
                Write-Host "    List of services which are already in running state : " -f Cyan -NoNewline
                Write-Host "NIL"
                "       List of services which are already in running state : NIL" >> $LogFile
                $fn_htmlBody += "<font size='2'><pre class='tab1'><b>   List of services which are already in running state : NIL</b></pre></font><br>"
            }
            if($result.Startservices -and $result.Unstartservices)
            {
                $partial_start_servers += $servername
            }
            if($result.Startservices)
            {
                $services_started += $servername
            }
            if($result.Unstartservices)
            {
                $services_not_started += $servername
            }
    
            # test the connection with the db instance to check if the services are started correctly
            $serverInstance = $servername
            $databaseName = "master"
            $connectionString = "Server=$serverInstance;Database=$databaseName;Integrated Security=True;"
            $query = "SELECT 1;"
            try
            {
                $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
                $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
                $connection.Open()
                $result = $command.ExecuteScalar()
                if($result -eq 1) 
                {
                    Write-Host "Connection Establishment to the db instance : " -nonewline
                    Write-Host "Success" -f Green
                    "   Connection Establishment to the db instance : Success" >> $LogFile
                    $fn_htmlBody += "<font size='2' color='green'><pre class='tab1'><b> Connection Establishment to the db instance : Success</b></pre></font><br>"
    
                } 
                else 
                {
                    Write-Host "Connection Establishment to the db instance : Failed"
                    "   Connection Establishment to the db instance : Failed" >> $LogFile
                    $fn_htmlBody += "<font size='2' color='red'><pre class='tab1'><b> Connection Establishment to the db instance : Failed</b></pre></font><br>"
                    $connection_failed += $servername
                }
            } 
            catch 
            {
                Write-Host "Error in establising the connection to the SQL instance: $_.Exception.Message"
                "   Error in establising the connection to the SQL instance: $_.Exception.Message" >> $LogFile
            } 
            finally 
            {
                if($connection.State -ne 'Closed') 
                {
                    $connection.Close()
                }
            }
            $end_result = [PSCustomObject]@{
                ExitCode = 0
                Content = $fn_htmlBody
                ConnectionFailed = $connection_failed
                ServicesStarted = $services_started
                ServicesNotStarted = $services_not_started
            }
            return $end_result 
    
        }
        else
        {
            Write-Host "There is no service listed to start"
            "There is no service to start" >> $LogFile
            $fn_htmlBody = "<font size='2' color='red'><pre class='tab1'><b> There is no service to start</b></pre></font><br>"
            $end_result = [PSCustomObject]@{
                ExitCode = -1
                Content = $fn_htmlBody
            }
            return $end_result 
        }

    }
    catch 
    {
        Write-Host "Error in establishing the connection to the server : $servername"
        "Error in establishing the connection to the server : $servername" >> $Logfile
        Write-Host "Error : $_"
        "Error : $_" >> $LogFile
        $fn_htmlBody += "<font size='2' color='red'><b>Error : $_</b></font><br>"
        $end_result = [PSCustomObject]@{
            ExitCode = 1
            Content = $fn_htmlBody
        }
        return $end_result
    }
}

function InsertValues
{
    param($username, $csvFilePath, $servers_record, $CRNumber, $timestamp)
    try
    {
        $TimeStamp = $timestamp
        $ExecuteBy = $username
        $BlackoutTime = ""
        $input_details = Import-CSV $csvFilePath

        $servers_list = $input_details.ServerName
        if($servers_list.count -eq 0)
        {
            break
        }
        foreach($server in $servers_list)
        {
            $EffectedServices = ($input_details | Where-Object {$_.ServerName -eq $server}).Services

            $query_to_insert = "INSERT INTO StopStartHistory (ServerName, BlackoutDuration, EffectedServices, TimeStamp, CRNumber, ExecuteBy, Job) VALUES ('$server', '$BlackoutTime', '$EffectedServices', '$TimeStamp', '$CRNumber', '$ExecuteBy', 'START');"
        
            Invoke-Sqlcmd -serverinstance "ITSUSRAWSP10439" -Database "StopStartTestDB" -query $query_to_insert
        
        }

    }
    catch
    {
        Write-Host "Error Occurred :  $_"
    }

}
try
{
    $time = Get-Date -format "yyyyMMddHHmm"
    $task = "START"
    $username = $($env:username)
    #---------------------------------Workfolder-----------------------------------
    $workfolder = "D:\StopStartServices"
    if(!(Test-Path $workfolder))
    {
        New-Item -Path $workfolder -ItemType Directory | Out-Null
    }
    
    #---------------------------------Logfolder and log file creation-----------------------------------
    $logFolder = "$workfolder\Logs"
    if(!(Test-Path $logFolder))
    {
        New-Item -Path $logFolder -ItemType Directory | Out-Null
    }
    else
    {
        $currentDate = Get-Date
        $daysToKeep = 60
        $dateThreshold = $currentDate.AddDays(-$daysToKeep)
        $filesToDelete = Get-ChildItem -Path $logFolder | Where-Object { $_.LastWriteTime -lt $dateThreshold }
        foreach ($file in $filesToDelete) {
            Remove-Item -Path $file.FullName -Force
        }
    }
    $global:LogFile = "$logFolder\Log_$($task)_$time.txt"
    if(!(Test-Path $LogFile))
    {
        New-Item -Path $LogFile -ItemType File | out-null
    }
    
    
    #---------------------------------declaration of variable to store htmlbody content--------------------------------------
    $global:htmlBody = "<html><head><style>
    body {
        font-family: Arial;
        }
    p  {
        color: black;
        font-family: Arial;
        }
      .tab{tab-size:4;}.tab1{tab-size:8;}.tab2{tab-size:16;}</style></head><body><br>"
    $global:htmlBody_mirror = $htmlBody   
    $delimiter = "********************************************************************************"
    #---------------------------------ctaskname--------------------------------------

    $csvoutputpath = "$workfolder\CSVOutput"
    Write-Host "`n$delimiter`n"
    $ctaskNumber = (Read-Host "Enter the CR number").trim()
    Write-Host "`n$delimiter`n"
    $global:Ctask = $ctaskNumber
    if($ctaskNumber)
    {
        $csvfiles = Get-ChildItem $csvoutputpath -Filter "*.csv"
        $iterate = 1
        $found = 1
        foreach($file in $csvfiles)
        {
            if(($file.Name).contains($ctaskNumber))
            {
                $filepath = $file.FullName
                $found = 0
                break
            }
            else
            {
                $iterate += 1
            }
        }
        if($found -eq 1 -and ($iterate-1) -eq $csvfiles.Count)
        {
            Write-Host "There is no suitable csv file found to start the services. Exiting..."
            "There is no suitable csv file found to start the services. Exiting..." >> $LogFile
            $htmlBody += "<font size='2'><b><font color='red'>There is no suitable csv file found to start the services. Exiting...</font></b></font><br>"
            TriggerMail $htmlBody
            Exit 1
        }
        
    }
    else 
    {
        Write-Host "There is no CTask number mentioned in the text file $csvtaskfilepath..Exiting.."
        "There is no CTask number mentioned in the text file $csvtaskfilepath..Exiting.." >> $LogFile
        $htmlBody += "<font size='2'><b><font color='red'>There is no CTask number mentioned in the text file $csvtaskfilepath..Exiting..</font></b></font><br>"
        TriggerMail $htmlBody
        Exit 1
    }
    
    #$filepath = "C:\StopStartServices\StopStartServices_STOP_202402281040.csv"
    $global:csvFilePath = "$csvoutputpath\StopStartServices_$($task)_$($time)_$($Ctask).csv"
    if(!(Test-Path -path $csvFilePath))
    {
        New-Item -path $csvFilePath -ItemType File | Out-Null
    }
    $servers_record = Import-CSV $filepath
    $servers_started = @()
    $servers_not_started = @()
    $servers_connection_failed = @()
    $timestamp = get-date -f "dd/MM/yyy HH:mm:ss"
        foreach($record in $servers_record)
        {
            $servername = $record.servername
            if($record.Services)
            {
                Write-Host "Server : $servername" -f Magenta
                "Server : $servername" >> $LogFile
                $htmlBody += "<font size='5'><b><u>Server : $servername</u></b></font><br>"
                $services_to_start = $record.Services
                $startservices_result = StartServices $servername $services_to_start
                if($startservices_result.ExitCode -eq 0)
                {
                    Write-Host "`n$delimiter`n"
                    "`n$delimiter`n" >> $LogFile
                    $htmlBody += $startservices_result.Content
                    $htmlBody += "<br>$delimiter<br>"
                    $servers_started += $startservices_result.ServicesStarted
                    $servers_not_started += $startservices_result.ServicesNotStarted
                    $servers_connection_failed += $startservices_result.ConnectionFailed
                }
            }
        }
        

        $detailed_HtmlBody = "<font size='4'><b><u><font color='DodgerBlue'>Summary</font></u></b></font>"
        if($servers_started)
        {
            Write-Host "The list of servers where the services are started successfully..." -f Green
            #"The list of servers where the services are started successfully...`n" >> $LogFile
            $detailed_HtmlBody += "<font size='3'><b><p>Services started <font color='Green'>Successfully</font></p></b></font>"
            foreach($server in $servers_started)
            {
                Write-Host $server
                #$server >> $LogFile
                $detailed_HtmlBody += "<font size='2'><pre class='tab1'>    <li>$server</li></font><br>"
            }
            
        }
        #"intermediate -------------------> $intermediate_array" >> $LogFile
        #"$($intermediate_array).gettype()" >> $LogFile
        #"servers_not_started ------------------------> $servers_not_started" >> $LogFile
        #"$servers_not_started.gettype()" >> $LogFile
        $intermediate_array = $servers_not_started
        $servers_not_started = $servers_started | ? {$intermediate_array -contains $_}
        #$servers_not_started.gettype() >> $LogFile
        if($servers_not_started)
        {
            $detailed_HtmlBody += "<br>$delimiter<br>"
            Write-Host "The list of servers where the services are not started successfully..." -f Green
            "The list of servers where the services are not started successfully...`n" >> $LogFile
            $detailed_HtmlBody += "<font size='3'><b><p>Services <font color='Red'>not started</font></p></b></font>"
            foreach($server in $servers_not_started)
            {
                Write-Host $server
                $server >> $LogFile
                $detailed_HtmlBody += "<font size='2'><pre class='tab1'>    <li>$server</li></font><br>"
            }
            $detailed_HtmlBody += "<br>$delimiter<br>"
        }  
        if($servers_connection_failed)
        {
            Write-Host "The list of servers where the connection establishment failed...`n" -f Red
            "The list of servers where the connection establishment failed...`n" >> $LogFile
            $detailed_HtmlBody += "<font size='3'><b><p><font color='Red'>Connection not established </font></p></b></font>"
            foreach($server in $servers_connection_failed)
            {
                Write-Host $server
                $server >> $LogFile
                $detailed_HtmlBody += "<font size='2'><pre class='tab1'>    <li>$server</li></font><br>"
            }
            $detailed_HtmlBody += "<br>$delimiter<br>"
        }  
    
        $detailed_HtmlBody += "<br><br><b><p>-----------------------------------------------------------------------Details logs are below---------------------------------------------------------------------------------</p></b><br>"
        $htmlBody += "<br><font size='3' color='green'><i>Kindly find the logs at $LogFile in the central server. Attached here for reference.</i></font></body></html>"
        $htmlBody = $htmlBody.Replace($htmlBody_mirror,"")
        $main_HtmlBody = $htmlBody_mirror + $detailed_HtmlBody + $htmlBody
    
        TriggerMail $main_HtmlBody $username

        Write-Host "`nKindly find the logs at $LogFile" -f Green
        $csvfolder = "$workfolder\CSVOutput"
        $files = Get-ChildItem -Path $csvfolder -File
        $latestFile = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $csvFilePath = $latestFile.FullName
        $filepath = "$workfolder\StopServices_Servers.csv"
        $servers_record = Import-CSV $filepath

        InsertValues $username $csvFilePath $servers_record $CRNumber $timestamp
        $refresh_query = "EXEC [SERVICES].[msdb].[dbo].[sp_start_job] @job_name = N'3A07479C-75FB-42AA-AED1-FB96D95C7E0D';"
        Invoke-Sqlcmd -ServerInstance "ITSUSRAWSP10439"-query $refresh_query


}

catch
{
    Write-Host "Error Occurred : $_"
}
