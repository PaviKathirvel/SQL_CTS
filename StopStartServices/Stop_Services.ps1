function TriggerMail
{
    param($htmlBody, $username)
    $username = $($env:username) -ireplace "Admin_",""
    $smtpServer = "smtp.na.jnj.com"
    $from = "SA-NCSUS-SQLHELPDESK@its.jnj.com"
    $to = "$($($username -ireplace '^Admin_',''))@its.jnj.com"
    $cc = "PKathirv@its.jnj.com", "DL-ITSUS-GSSQL@ITS.JNJ.COM"
    $subject = "Log Report for the task Stop Services : $CRNumber"
    $body = $htmlBody
    $attachment = $Logfile

    Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -CC $cc -Subject $subject -Body $body -Attachments $attachment -BodyAsHtml -UseSsl 
}
function GetCRDetails
{
    # Connects to the IRIS through API call and collects the mentioned CR details
    param
    (
        [string] $CR,
        [string] $Debug
    )

    #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
    #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Ssl3
    
    #[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

    $Type = "application/json"
    $method = "post"
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add('Content-Type','application/x-www-form-urlencoded')

    # Specify endpoint uri
    $uri = "https://login.microsoftonline.com/its.jnj.com/oauth2/token?api-version=1.0"

    # Specify HTTP method
    $BaseURL = "https://jnj-internal-production.apigee.net/apg-001-servicenow/v1/now"
    $IRIS_APIs = "NKx8Q~6pmejthqqiIrEytC5evEQlecCcbKr9Ea39"
    if($IRIS_APIs)
    {
    	$bodyJson = "grant_type=client_credentials&client_id=edaff30a-eb13-441a-a9ca-830bc31c165b&client_secret=$($IRIS_APIs)&resource=https%3A//ITS-APP-ISM-IRIS-Prod.jnj.com"
    }
    else
    {
    	exit
    }


    # Send HTTP request
    $response = Invoke-WebRequest -Headers $headers -Method $method -Uri $uri -Body $bodyJson | ConvertFrom-Json | select access_token, token_type

    if ($Debug -eq "Y")
    {
    	LogActivity $response.access_token
    }
    $SNOWSessionHeader = @{'Authorization' = "$($response.token_type) $($response.access_token)"}

    #Setup file/folder for excel report
    $CurrDir = Get-Location
    Write-Output $CurrDir.Path
    $OutputPath = $CurrDir.Path -replace "PS", "PS_Reports"
    if ((Test-Path -path $OutputPath) -ne $True)
    {
    	New-Item $OutputPath -type directory
    }

    ###############################################
    # Getting Database CI/attrinute Info
    ###############################################
    try
    {
        $CHGURL = "$($BaseURL)/table/change_request?sysparm_query=number=$($CR)&sysparm_display_value=true"
        $CHGJSON = Invoke-RestMethod -Method GET -Uri $CHGURL -TimeoutSec 100 -Headers $SNOWSessionHeader -ContentType $Type  -Proxy $null 
        $CHG_RESULTS = $CHGJSON.result
    }
    catch
    {
        Write-Host "$_"
        Write-Host "`nNo CR found in IRIS. Please re-verify`n" -f Red
        Exit 0
        $htmlBody += "<font>Error occurred in fetching the data.`nPlease verify the correctness of the entered CR number.</font>"
    }
    if($CHGJSON.result.count -eq 0)
    {
        Write-Host "`nNo CR found in IRIS. Please re-verify`n" -f Red   
        $htmlBody += "<font>No data fetched from the CR API call. Please verify the correctness of the CR number.</font>"
        return $null
    }
    if($CHG_RESULTS.state -ne 'Implement')
    {
        $startdate = 0 
        $enddate = 0
    }
    else
    {
        $startdate = Get-Date $($CHG_RESULTS.work_start | Select-Object -Unique)
        $enddate = Get-Date $($CHG_RESULTS.end_date | Select-Object -Unique)
    }
    $result = [PSCustomObject]@{
        State = $CHG_RESULTS.state | Select-Object -Unique
	    StartDate = $startdate
	    EndDate = $enddate
        AffectedCIs = $($CHG_RESULTS.u_all_ci | Select-Object -Unique) -split ";"
    }
    return $result
}
function ValidateCR
{
    param($state, $startdate, $enddate, $affectedCIs,
          [string[]]$servers
        )
   
    if(!($state -eq "Implement"))
    {
        Write-Host "The CR is not in implement state. Exiting" -f Red
        $htmlBody += "<font size='2'><b><font color='red'>The CR is not in implement state. Exiting</font></b></font>"
        TriggerMail $htmlBody
        Exit 1
    }
    Write-Host "   CR is in implement state? Yes" -f Cyan

    $currentDate = Get-Date 
    if(!($startdate-le $currentDate -and $enddate -ge $currentDate))
    {
        Write-Host "`nThe current time does not fall in between the schedule window of the CR. Exiting" -f Red
        $htmlBody += "<font size='2'><b><font color='red'>The current time does not fall in between the schedule window of the CR. Exiting</font></b></font>"
        Exit 1
    }
    Write-Host "   The execution time falls in between the schedule window? Yes" -f Cyan
    $servers = $servers | select-object -unique | % {$_.trim()} | % {$_.ToUpper()}
    if(($servers | Where-Object {$affectedCIs -notcontains $_}))
    {
        Write-Warning "The CSV file has extra servers comparing with the affectedCIs of the given CR" #`nAffected CIs : $($affectedCIs -join ',')"
        #Write-Host "Extra Servers : `n$($($servers | Where-Object {$affectedCIs -notcontains $_}) | ForEach-Object {"$_"} -join '  ```n')"
        Write-Host "Unmatched Servers in CSV File : " -NoNewLine
        Write-Host "$($($servers | Where-Object { $affectedCIs -notcontains $_ } | ForEach-Object { $_ }) -join '  ')" -f Red
        Write-Host "Matched Servers in CSV File : " -NoNewline
        Write-Host "$($($servers | Where-Object {$affectedCIs -contains $_}) -join '  ')" -f Green
        $prompt = Read-Host "`nGood to procceed further Y/N?"
        if($prompt -ieq "y" -or $prompt -ieq "Y")
        {
            Write-Host "`nProceeding.." -f Green
        }
        else 
        {
            Write-Host "`nExiting.." -f Red
            Exit 1
        }
    }
    else 
    {
        Write-Host "   CSV File Servers are matched with the Affected CIs? Yes" -f Cyan
    }
    Write-Host "`nCR Validation is completed successfully`n" -f Green
    return $($($servers | Where-Object {$affectedCIs -contains $_}))
}
function Blackout
{
    param($servers_list, $BlackoutTime_list)
    Write-Host "`n"
    $successful_blackout = @()
    $failed_blackout = @()

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    
    foreach($servername in $servers_list)
    {
        $BlackoutTime = $BlackoutTime_list
        [int]$DurationHours = $BlackoutTime.Split(":")[0]
        [int]$DurationMins = $BlackoutTime.split(":")[1]
        $duration = 0
        $duration += $DurationHours * 60 * 60
        $duration += $DurationMins * 60

        $moogsoftUrl = "https://sddc-moogq-ui1.jnj.com/graze/v1/createMaintenanceWindow"
        $username = "grazedbteam"
        $encryptedString = get-content -path "D:\StopStartServices\Debug\VIA_Powershell\pwd_encrypted.txt"
        $secureString = ConvertTo-SecureString $encryptedString -Key @(1..16)
        $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString))
        $credentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($username):$($password)"))
        $headers = @{
            Authorization = "Basic $credentials"
            "Content-Type" = "application/json; charset=UTF-8"
        }
        $startDateTimeEpoch = [int][double]::Parse((Get-Date (Get-Date).ToUniversalTime() -UFormat %s))
        $body_win_server = @"
        {
            "name": "Stop-Start Services",
            "description": "Blackout for the server $servername on behalf of the CR $CRNumber",
            "filter": "source = \"$servername\" and class = \"cmdb_ci_win_server\"",
            "start_date_time": $startDateTimeEpoch,
            "duration": $duration,
            "forward_alerts": false,
            "timezone": "America/New_York"
        }
"@
        $body_db_server = @"
        {
            "name": "Stop-Start Services",
            "description": "Blackout for the server $servername on behalf of the CR $CRNumber",
            "filter": "source = \"$servername\" and class = \"cmdb_ci_db_mssql_instance\"",
            "start_date_time": $startDateTimeEpoch,
            "duration": $duration,
            "forward_alerts": false,
            "timezone": "America/New_York"
        }
"@

        try 
        {
            $response_win = Invoke-RestMethod -Uri $moogsoftUrl -Method Post -Headers $headers -Body $body_win_server
            $response_db = Invoke-RestMethod -Uri $moogsoftUrl -Method Post -Headers $headers -Body $body_db_server
            $successful_blackout += $servername  
            Write-Host "The blackout for the server $servername is successful`n" -ForegroundColor Green
            $response_win | ConvertTo-Json
            $response_db | ConvertTo-Json
        } 
        catch 
        {
            Write-Host "Error Occurred in blackout the server : $servername" -ForegroundColor Red
            Write-Host "    Error :  $_`n"
            "Error Occurred in blackout the server : $servername" >> $Logfile
            #"    Error :  $_`n" >> $LogFile
            $failed_blackout += $servername        
        }

    }

    $returnobject = [PSCustomObject]@{
        SuccessfulBlackout = $successful_blackout
        FailedBlackout = $failed_blackout
    }
    return $returnobject
}
function StopServices
{
    param(
        $servername,
        $csvFilePath
    )
    try
    {

        $result = Invoke-Command -ComputerName $servername -ScriptBlock{
            $is_cluster = Get-Service | Where-Object {$_ -like "Cluster"}
                if($is_cluster)
                {
                    return "cluster"
                }
                $sqlservices = Get-Service | Where-Object {$_.DisplayName -like "*SQL*"}
                if($sqlservices)
                {
                    $totalservices = $sqlservices.Name
                    $before_services_info = New-Object System.Collections.ArrayList
                    foreach($service in $totalservices)
                    {
                        $before_serviceinfo = Get-Service $service | select-object DisplayName, status, Starttype
                        $before_services_info += $before_serviceinfo
                    }
                    $sqlagentservice = Get-Service | Where-Object {$_.DisplayName -ilike "*SQL*Server*" -and  $_.Name -ilike "*SQL*Server*"}
                    if($sqlagentservice.Status -eq "Running")
                    {
                        $is_alwayson = (Invoke-SqlCmd -ServerInstance $server -Database "master" -Query "SELECT SERVERPROPERTY('IsHadrEnabled') as is_hadr_enabled" ).is_hadr_enabled
                        if($is_alwayson -eq 1)
                        {
                            return "alwayson"
                        }
                    }
                    else 
                    {
                        $resultObject = [PSCustomObject]@{
                            Status = "sqlagent_stopped"
                            BeforeServices_Info = $before_services_info               
                        }
                        return $resultObject
                    }
                }
                $upservices = Get-Service | Where-Object {$_.DisplayName -like "*SQL*"} | Where-Object {$_.Status -eq "Running"}
                $services_to_stop = $upservices.Name
                if($upservices)
                {
                    $timestamp = Get-Date
                    $timezone = (Get-TimeZone).id
                    $stoppableservices = @()
                    $unstoppableservices = @()
                    $services_info = New-Object System.Collections.ArrayList
                    foreach($service in $services_to_stop)
                    {
                        try
                        {
                            Set-Service $service -StartupType Manual -ErrorAction Stop
                            Stop-Service $service -force -ErrorAction Stop
                            $stoppableservices += $service
                            $info = get-service $service 
			                $services_info += $info
                        }
                        catch 
                        {
                            $unstoppableservices += $service
                        } 
                    }
                    $resultObject = [PSCustomObject]@{
			            Services_Info = $services_info
                        BeforeServices_Info = $before_services_info
                        TotalServices = $services_to_stop
                        StopServices = $stoppableservices
                        UnStopServices = $unstoppableservices
                        Timestamp = $timestamp       
                        TimeZone = $timezone                
                    }
                    return $resultObject
                }
                else 
                {
                    return $null
                }
                
        }
        if($result)
        {
            if(!($result.gettype().Name -ieq "String") -and ($result.Status -eq $null))
            {
                Write-Host "Connection To Windows : Established"
                "   Connection To Windows: Established" >> $LogFile
                $fn_htmlBody += "<font size='2'><pre class='tab'>   Connection To Windows         : <font color='green'>Established</font></pre></font>"
                
                Write-Host "Connection to SQL instance :  Established"
                "   Connection to SQL instance :  Established" >> $LogFile
                $fn_htmlBody += "<font size='2'><pre class='tab'>   Connection to SQL instance    : <font color='green'>Established</font></pre></font>"
                
                Write-Host "ServerTime : $($result.Timestamp) $($result.TimeZone)"
                "ServerTime : $($result.Timestamp) ($($result.TimeZone))" >> $LogFile
                $fn_htmlBody += "<font size='2'><br><pre class='tab'>   ServerTime : $($result.Timestamp) ($($result.TimeZone))</pre></font><br>"
                "`n   The status of the services before stopping the services" >> $LogFile
                $($result.BeforeServices_Info) | Select-Object DisplayName, Status, StartType >> $LogFile

                "`n   The status of the services after stopping the services are below" >> $LogFile                
                $($result.Services_Info) | Select-Object DisplayName, Status, StartType >> $LogFile
            
                #$fn_htmlBody += "<font size='2'><pre class='tab'>   The status of the services are below : </pre></font>"
                $printresulthtml = $($result.Services_Info) | Select-Object DisplayName, Status, StartType
                $fn_htmlBody += "<font size='2'><pre class='tab'>   $printresulthtml</pre></font>"

                "   Categories : " >> $LogFile
                $fn_htmlBody += "<font size='2'><pre class='tab'><b>    Categories : <b></pre></font><br>"
                if($result.TotalServices)
                {
                    Write-Host "    List of services to be stopped : " -f Cyan
                    "       List of services to be stopped : " >> $LogFile
                    $fn_htmlBody += "<font size='2'><pre class='tab1'><b>   List of services to be stopped : </b></pre></font>"
                    $index = 1
                    foreach($service in $($result.TotalServices))
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
                    Write-Host "    List of services to be stopped : " -f Cyan -NoNewline
                    Write-Host "NIL" 
                    "       List of services to be stopped : NIL" >> $LogFile
                    $fn_htmlBody += "<font size='2'><pre class='tab1'><b>   List of services to be stopped : NIL</b></pre></font><br>"
                }
            
                if($result.StopServices)
                {
                    Write-Host "    List of services which are stopped : " -f Cyan
                    "       List of services which are stopped : " >> $LogFile
                    $fn_htmlBody += "<font size='2'><br><pre class='tab1'><b>   List of services which are stopped : </b></pre></font>"
                    $index = 1
                    foreach($service in $($result.StopServices))
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
                    Write-Host "    List of services which are stopped : " -f cyan -NoNewline
                    Write-Host "NIL"
                    "           List of services which are stopped : NIL" >> $LogFile
                    $fn_htmlBody += "<font size='2' color='red'><pre class='tab1'><b>   List of services which are stopped : NIL</b></pre></font><br>"
                }
            
                if($($result.UnStopServices))
                {
                    Write-Host "     List of services which are not stopped : " -f Cyan
                    "       List of services which are not stopped : " >> $LogFile
                    $fn_htmlBody += "<font size='2' color='red'><pre class='tab1'><b>   List of services which are not stopped : </b></pre></font><br>"
                    $index = 1
                    foreach($service in $($result.UnStopServices))
                    {
                        Write-Host "        $index.$service"
                        "           $index.$service" >> $LogFile
                        $fn_htmlBody += "<font size='2'><pre class='tab2'>      $index.$service</pre></font>"
                        $index++
                    }
                    $fn_htmlBody += "<br>"
                    $partially_failed_servers += $servername
                }
                else
                {
                    Write-Host "     List of services which are not stopped : " -f Cyan -NoNewline
                    Write-Host "NIL"
                    "    List of services which are not stopped : NIL" >> $LogFile
                    $fn_htmlBody += "<font size='2'><pre class='tab1'><b>   List of services which are not stopped : NIL</b></pre></font><br>"
                }
                $exitcode = 0
                
            }
            elseif($result -ieq "cluster")
            {
                $exitcode = 2
                Write-Host "The server is a clustered server.`nPlease proceed manually to stop the services" -f Yellow
                "The server is a clustered server.`nPlease proceed manually to stop the services" >> $LogFile
                $fn_htmlBody += "<font size='2'><br>The server is a clustered server. Please proceed manually to stop the services</font>"
            }
            elseif($result -ieq "alwayson")
            {
                #$alwayson_servers += $servername
                $exitcode = 3
                Write-Host "The server is an alwayson server.`nPlease proceed manually to stop the services" -f Yellow
                "The server is an alwayson server.`nPlease proceed manually to stop the services" >> $Logfile
                $fn_htmlBody += "<font size='2'><br>The server is an alwayson server. Please proceed manually to stop the services</font>"
            }
            elseif($result.Status -ieq "sqlagent_stopped")
            {
                #$already_stopped_servers += $servername
                $exitcode = 4
                Write-Host "The service MSSQLSERVER is already in stopped state.`nPlease proceed manually to check the status of other services" -f Yellow
                "The service MSSQLSERVER is already in stopped state.`nPlease proceed manually to check the status of other services" >> $Logfile
                "`n The status of the services are below : " >> $LogFile
                $($result.BeforeServices_Info) | Select-Object DisplayName, Status, StartType >> $LogFile
                $fn_htmlBody += "<font size='2'><br>The service MSSQLSERVER is already in stopped state. Please proceed manually to check the status of other services<br></font>"

            }
        }
        else
        {
            "The services are already in stopped state. No services to be stopped" >> $LogFile
            $fn_htmlBody += "<font size='2' color='green'><br>The services are already in stopped state. No services to be stopped</font><br>"
            $end_result = [PSCustomObject]@{
                ExitCode = -1
                Content = $fn_htmlBody
            }
            return $end_result
        }
        $data = @()
	    $updateservices = $($result.StopServices) -join ","
	    $record = [PSCustomObject]@{
		    ServerName = $servername
		    Services = $updateservices
	    }
	    $data += $record
	    $data | Export-Csv -Path $csvFilePath -Append -NoTypeInformation
        $end_result = [PSCustomObject]@{
            ExitCode = $exitcode
            Content = $fn_htmlBody
        }
        return $end_result 
    }
    catch
    {
        Write-Host "Error in establishing the connection to the server : $servername" -f Red
        "Error in establishing the connection to the server : $servername" >> $Logfile
        $fn_htmlBody += "<font size='2' color='red'><b>Error in establishing the connection to the server : $servername</b></font><br>"
        Write-Host "Error : $_"
        "Error : $_" >> $LogFile
        $fn_htmlBody += "<font size='2' color='red'><b>Error : $_</b></font><br>"
        $end_result = [PSCustomObject]@{
            ExitCode = $exitcode
            Content = $fn_htmlBody
        }
    }
}
function InsertValues
{
    param($username, $csvFilePath, $servers_record, $CRNumber, $timestamp, $BlackoutTime)
    try
    {
        $TimeStamp = $timestamp
        $ExecuteBy = $username
    
        $input_details = Import-CSV $csvFilePath

        $servers_list = $input_details.ServerName
        foreach($server in $servers_list)
        {
            $EffectedServices = ($input_details | Where-Object {$_.ServerName -eq $server}).Services

            $query_to_insert = "INSERT INTO StopStartHistory (ServerName, BlackoutDuration, EffectedServices, TimeStamp, CRNumber, ExecuteBy, Job) VALUES ('$server', '$BlackoutTime', '$EffectedServices', '$TimeStamp', '$CRNumber', '$ExecuteBy', 'STOP');"
        
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
    $task = "STOP"
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
    $delimiter = "*******************************************************************************************************************"
    #---------------------------------CSV folder and the csv file creation--------------------------------------
    $csvfolder = "$workfolder\CSVOutput"
    if(!(Test-Path -path $csvfolder))
    {
        New-Item -path $csvfolder -ItemType Directory | out-null
    }
    else 
    {
        $currentDate = Get-Date
        $daysToKeep = 60
        $dateThreshold = $currentDate.AddDays(-$daysToKeep)
        $filesToDelete = Get-ChildItem -Path $csvfolder | Where-Object { $_.LastWriteTime -lt $dateThreshold }
        foreach ($file in $filesToDelete) {
            Remove-Item -Path $file.FullName -Force
        }
    }
    $filepath = "$workfolder\StopServices_Servers.csv"
    $servers_record = Import-CSV $filepath
    Write-Host "`n$delimiter`n"
    $CRNumber = Read-Host "Enter the CR Number" #"CHG000011196431" #
    Write-Host "`n$delimiter`n"
    $servers = $servers_record.ServerName
    if(!($CRNumber))
    {
        Write-Host "Please provide the valid CR number. Exiting.." -f Red
        "Please provide the CR number. Exiting" >> $LogFile
        $htmlBody += "<font size='2'><b><font color='red'>Please provide the CR number. Exiting..</font></b></font><br>"
        TriggerMail $htmlBody
        Exit 1
    }

    [PSCustomObject]$result = GetCRDetails -CR $CRNumber
    if(!($result[-1]))
    {
        Write-Host "Error occurred in fetching the data through API call. Kindly make sure the CRNumber is valid" -f Red
        $htmlBody += "<font size='2'><b><font color='red'>Error occurred in fetching the data through API call. Kindly make sure the CRNumber is valid</font></b></font>"
        TriggerMail $htmlBody
        Exit 1
    }

    Write-Host "Validating the CR details..."
    $considered_servers = ValidateCR $result.State $result.StartDate $result.EndDate $result.AffectedCIs $servers
    Write-Host "$delimiter`n"
    $global:csvFilePath = "$csvfolder\StopStartServices_$($task)_$($time)_$($CRNumber).csv"
    if(!(Test-Path -path $csvFilePath))
    {
        $csvfiles = Get-ChildItem $csvfolder -Filter "*.csv"
        foreach($file in $csvfiles)
        {
            if(($file.Name).contains("$CRNumber"))
            {
                Write-Host "The job has already created a file with the CRNumber $CRNumber. Cannot run the job again" -f Red
                "The job has already created a file with the CRNumber $CRNumber. Cannot run the job again" >> $LogFile
                $htmlBody += "<font size='2'><b><font color='red'>The job has already created a file with the CRNumber $CRNumber. Cannot run the job again</font></b></font><br>"
                TriggerMail $htmlBody
                Exit 1
            }
        }
        New-Item -path $csvFilePath -ItemType File | Out-Null
    }

    #-----------------------------------------------------------------------------------------------------------------------------------------------#
    $blackout_duration = "1:00"
    Write-Host "Blackout Duration Details..."
    Write-Host "   The default blackout duration will be " -NoNewline
    Write-Host "1 hr." -f Magenta
    $blackout_prompt = Read-Host "   Would you like to change the default blackout duration? Provide Y/N "
    if($blackout_prompt -ieq 'y')
    {
        Write-Host "   Provide the blackout duration time in the format " -NoNewline; Write-Host "'HH:mm' " -f Magenta -NoNewline; Write-Host ": " -NoNewline
        $blackout_duration = Read-Host
        while(!($blackout_duration.contains(":")))
        {
            $blackout_duration = Read-Host "   The given time format is invalid. Please provide in the format 'HH:mm' "
        }
    }
    Write-Host "`nThe blackout duration is " -NoNewline
    Write-Host "'$blackout_duration'" -f Magenta
    #Write-Host "`n$delimiter`n"
    #---------------Collect Production Servers List from the central server and check for prod and non-prod servers in the given input list------------------
    $production_servers = @((Invoke-SqlCmd -ServerInstance $(hostname) -Database "sqlmon" -Query "select * from [dbo].[tbl_serverlist] where ServerType = 'production'").ServerName)
    $non_prod_servers = $considered_servers | Where-Object {$production_servers -notcontains $_}
    $prod_servers = $considered_servers | Where-Object {$production_servers -contains $_}

    #-----------------Call Blackout function to blackout all the given servers ---------------------------------------#
    $blackout_detail = Blackout $($non_prod_servers) $($blackout_duration)
    $success_blackout_servers = $($blackout_detail.SuccessfulBlackout)
    $failed_blackout_servers = $($blackout_detail.FailedBlackout)

    Write-Host "Blackout Buffer Time. Sleeping for 15 Seconds...`n"
    "Blackout Buffer Time. Sleeping for 15 Seconds..." >> $LogFile
    Start-Sleep -s 15
    
    Write-Host "$delimiter`n"

    #-----------collect the list of servers where services stopped successfully or any error------------
    $global:clustered_servers = @()
    $global:alwayson_servers = @()
    $global:success_servers = @()
    $global:already_stopped_servers = @()
    $global:connection_failed_servers = @()
    $global:partially_failed_servers = @()
    $timestamp = get-date -f "dd/MM/yyyy HH:mm:ss"
    foreach($servername in $success_blackout_servers)
    {
        #$servername = $record.servername
        if($servername)
        {
            #$duration = $servers_record.DurationTime[$servers_record.ServerName.indexof($servername)]
            Write-Host "Server : $servername" -f Magenta
            "Server : $servername" >> $LogFile
            $htmlBody += "<font size='5'><b><u>Server : $servername</u></b></font><br>"
            #$blackout_result = Blackout $servername $duration
            #if($blackout_result[-1] -eq 0)
            #{
                Write-Host "    Blackout : Done"
                "   Blackout : Done" >> $LogFile
                $htmlBody += "<font size='2'><b><br><pre class='tab'>   Blackout                      : <font color='green'>DONE</font></pre></b></font>"
                $stopservices_result = StopServices $servername $csvFilePath
                if($stopservices_result.ExitCode -eq 0)
                {
                    Write-Host "`nServices stopped" -f Green
                    #Write-Host "Data exported to $csvFilePath"
                    "`nServices stopped" >> $LogFile
                    "Data added to $csvFilePath" >> $LogFile
                    Write-Host "`n$delimiter`n"
                    "`n$delimiter`n" >> $LogFile
                    $htmlBody += $stopservices_result.Content
                    $htmlBody += "<br>$delimiter<br>"
                    $success_servers += $servername
                    $success_servers += "," 
                }
                elseif($stopservices_result.ExitCode -eq -1)
                {
                    Write-Host "`n$delimiter`n"
                    "`n$delimiter`n" >> $LogFile
                    $htmlBody += $stopservices_result.Content
                    $htmlBody += "<br>$delimiter<br>"
                    $already_stopped_servers += $servername
                    $already_stopped_servers += ""
                }
                elseif($stopservices_result.ExitCode -eq 2)
                {
                    Write-Host "`n$delimiter`n"
                    "`n$delimiter`n" >> $LogFile
                    $clustered_servers += $servername
                    $clustered_servers += ","
                    $htmlBody += $stopservices_result.Content
                    $htmlBody += "<br>$delimiter<br>"
                }
                elseif($stopservices_result.ExitCode -eq 3)
                {
                    Write-Host "`n$delimiter`n"
                    "`n$delimiter`n" >> $LogFile
                    $alwayson_servers += $servername
                    $alwayson_servers += ","
                    $htmlBody += $stopservices_result.Content
                    $htmlBody += "<br>$delimiter<br>"
                }
                elseif($stopservices_result.ExitCode -eq 4)
                {
                    Write-Host "`n$delimiter`n"
                    "`n$delimiter`n" >> $LogFile
                    $already_stopped_servers += $servername
                    $already_stopped_servers += ","
                    $htmlBody += $stopservices_result.Content
                    $htmlBody += "<br>$delimiter<br>"
                }
                else
                {
                    Write-Host "`nConnection: Error in connection establishment" -f Red
                    "`nConnection : Error in connection establishment" >> $LogFile
                    $htmlBody += "<br><font color='red'>Connection : Error in connection establishment</font><br>"
                    Write-Host "`n$delimiter`n"
                    "`n$delimiter`n" >> $LogFile
                    $htmlBody += "<br>$delimiter<br>"
                    $connection_failed_servers += $servername
                    $connection_failed_servers += ","
                }
        }
        #else 
        #{
        #    Write-Host "Blackout : Failed"
        #    "   Blackout : Failed" >> $LogFile
        #    $htmlBody += "<font size='2'><b><br><pre class='tab'>   Blackout                      : <font color='red'>FAILED</font></pre></b></font>"
        #    $htmlBody += "<br>*******************************************************************************<br>"
        #}
    }

        
    #}

    $detailed_HtmlBody = "<font size='4'><b><u><font color='DodgerBlue'>Summary</font></u></b></font>"

    $intermediate_success_log_file = "$workfolder\intermediate_success_log_file.txt"
    if(!(Test-Path -path $intermediate_success_log_file))
    {
        New-Item -path $intermediate_success_log_file -ItemType File | Out-Null
    }
    else
    {
        Remove-Item $intermediate_success_log_file -force | Out-Null
    }
    $intermediate_failed_log_file = "$workfolder\intermediate_failed_log_file.txt"
    if(!(Test-Path -path $intermediate_failed_log_file))
    {
        New-Item -path $intermediate_failed_log_file -ItemType File | Out-Null
    }
    else
    {
        Remove-Item $intermediate_failed_log_file -force | Out-Null
    }

    if($success_servers)
    {
        $value_string = $success_servers.split(",")
        foreach($var in $value_string)
        {
            #Write-Host "$var : $key"
            if(-not $var -or $var.Trim().Length -eq 0)
            {
                #"printing var ----------->$var" >> $intermediate_success_log_file   
            }
            else
            {
                "$var" >> $intermediate_success_log_file
            }
        }
        #$success_servers >> $intermediate_success_log_file
        
        #$success_servers = $success_servers.split(":")
        #Write-Host "The list of servers where the services are stopped successfully...`n"
        #"The list of servers where the services are stopped successfully...`n" >> $LogFile
        #$detailed_HtmlBody += "<font size='3'><b><p>Services Stopped <font color='Green'>Successfully</font></p></b></font>"
        #foreach($server in $success_servers)
        #{
        #    "Here I am $server" >> $LogFile
        #    "******************************************$success_servers" >> $LogFile
        #    "$($success_servers.gettype())" >> $LogFile
        #    Write-Host $server
        #    $server >> $LogFile
        #    $detailed_HtmlBody += "<font size='2'><pre class='tab1'>    <li>$server</li></font><br>"
        #}
        #Write-Host "`n********************************************************************************`n"
        #"`n********************************************************************************`n" >> $LogFile
        #$detailed_HtmlBody += "<br>********************************************************************************<br>"
    }    

    if($already_stopped_servers -or $prod_servers -or $clustered_servers -or $alwayson_servers -or $failed_blackout_servers -or $partially_failed_servers)
    {
        $total_failed_servers = @{}
        #Write-Host "Inside the hashtable loop"
        $total_failed_servers["Blackout Failed"] = $failed_blackout_servers
        $total_failed_servers["Prod Server"] = $prod_servers
        $total_failed_servers["Clustered Server"] = $clustered_servers
        $total_failed_servers["AlwaysOn Server"] = $alwayson_servers
        $total_failed_servers["Already services are in stopped state"] = $already_stopped_servers
        $total_failed_servers["Services are stopped partially"] = $partially_failed_servers

        foreach($key in $($total_failed_servers.Keys))
        {
            if($total_failed_servers[$key])
            {
                #$total_failed_servers[$key] = ($total_failed_servers[$key].split(":"))
                
                foreach($value in $($total_failed_servers[$key]))
                {
                    $value_string = $value.split(",")
                    #$value_string
                    foreach($var in $value_string)
                    {
                        #Write-Host "$var : $key"
                        if(-not $var -or $var.Trim().Length -eq 0)
                        {
                            #"printing var ----------->$var" >> $intermediate_failed_log_file   
                        }
                        else
                        {
                            "$var : $key" >> $intermediate_failed_log_file
                        }
                    }
                }
            }
        }

        #$detailed_HtmlBody += "<font size='2'><mark><br>Proceed manually to stop the services</mark></font><br>"
    }
    #if($connection_failed_servers)
    #{
    #    Write-Host "The list of servers where the connection establishment failed...`nPlease proceed manually on these servers`n"
    #    "The list of servers where the connection establishment failed...`nPlease proceed manually on these servers`n" >> $LogFile
    #    $detailed_HtmlBody += "<font size='2'><b>The list of servers where the connection establishment failed...`nPlease proceed manually on these servers</b></font><br>"
    #    foreach($server in $connection_failed_servers)
    #    {
    #        Write-Host $server
    #        $server >> $LogFile
    #        $detailed_HtmlBody += "<font size='2'><pre class='tab'> $server</font><br>"
    #    }
    #}

    $triumph = get-content $intermediate_success_log_file
    $defeat = get-content $intermediate_failed_log_file
    if($triumph)
    {
        #Write-Host "The list of servers where the services are stopped successfully...`n"
        #"The list of servers where the services are stopped successfully...`n" >> $LogFile
        $detailed_HtmlBody += "<font size='3'><b><p>Services Stopped <font color='Green'>Successfully</font></p></b></font>"
        Write-Host "Services stopped successfully on the below servers" -f Green
        foreach($line in $triumph)
        {
            Write-Host $line -f Cyan
            $detailed_HtmlBody += "<font size='2'><p>$line`n</p></font>"
        }
        
        $detailed_HtmlBody += "<br>$delimiter<br>"
    }
    if($defeat)
    {
        #Write-Host "The list of servers where the services are failed to stop..."
        #"The list of servers where the services are failed to stop..." >> $LogFile
        $detailed_HtmlBody += "<font size='3'><b><p>Services Stop <font color='Red'>Failed</font></p></b></font>"
        Write-Host "`n"
        Write-Warning "Services are not stopped on the below servers"
        foreach($line in $defeat)
        {
            $detailed_HtmlBody += "<font size='2'><p>$line`n</p></font>"
            Write-Host "$line" -f Yellow
        }
        #$detailed_HtmlBody += "<br>********************************************************************************************<br>"
    }

    Write-Host "`nData exported to $csvFilePath" -f Green
    "Data exported to $csvFilePath" >> $LogFile
    $htmlBody += "<br><font size='3'><i>Data exported to $csvFilePath</i></font>"
    Write-Host "Kindly find the logs at $LogFile`n" -f Green
    $htmlBody += "<br><font size='3'><i>Kindly find the logs at $LogFile in the central server. Attached here for reference.</i></font></body></html>"

    $detailed_HtmlBody += "<br><br><b><p>-----------------------------------------------------------------------Details logs are below---------------------------------------------------------------------------------</p></b><br>"

    $htmlBody = $htmlBody.Replace($htmlBody_mirror,"")
    $main_HtmlBody = $htmlBody_mirror + $detailed_HtmlBody + $htmlBody

    TriggerMail $main_HtmlBody $username
    #$main_HtmlBody >> $LogFile

    Remove-Item $intermediate_success_log_file -force
    Remove-Item $intermediate_failed_log_file -force
    #Write-Host "`nKindly find the logs at $LogFile" -f Green

    $files = Get-ChildItem -Path $csvfolder -File
    $latestFile = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $csvFilePath = $latestFile.FullName
    $filepath = "$workfolder\StopServices_Servers.csv"
    $servers_record = Import-CSV $filepath

    InsertValues $username $csvFilePath $servers_record $CRNumber $timestamp $blackout_duration
    $refresh_query = "EXEC [SERVICES].[msdb].[dbo].[sp_start_job] @job_name = N'3A07479C-75FB-42AA-AED1-FB96D95C7E0D';"
    Invoke-Sqlcmd -ServerInstance "ITSUSRAWSP10439"-query $refresh_query
}
catch
{
    Write-Host "Error Occurred : $_"
}