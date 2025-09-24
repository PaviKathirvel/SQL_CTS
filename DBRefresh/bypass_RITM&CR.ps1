$scriptName = "Auto_DBRefresh.ps1"
$scriptversion = 1.0
<#
SYNOPSIS
    Restores the databases in the SQL servers present across the regions.

SYNTAX
    .\Auto_DBRefresh.ps1 

DESCRIPTION
    This script takes four inputs from the user :  RITM number, CR number, Source DB(s), Target DB(s)
    Validates the correctness of the entered RITM and CR numbers. Dynamically picks the value of source and target servers from the RITM.
    Validates the servers accessibility and compatibility.
    Validates the source and target DB(s) with list of specifications.
    Check the space availability in the target server.
    Finds the nearest NAS location and saves the full backups of both source and target DB(s). 
    Saves the login scripts and properties of target DB(s) in the NAS location.
    Executes the restore of source DB(s) to target DB(s) respectively. Post retsore, logins & permissions will be re-applied.
NOTES
    Ensure that the script is run with administrative privileges.
    This version of script is not handling the restore of DB(s) with multiple data/log files
#>

###############################################################################
#   Version		    Date		        Author				Modification    
#   1.0             15 Dec 2024         Pavithra            New Script is build
###############################################################################
function LogActivity 
{
    #Writes the logs to the console and logfile
    param($content, $color, $htmlenabled=0, $console_disabled=1, $log_disabled=0)
    try
    {
        if(!$color)
        {
            $color = "White"
        }
        if($htmlenabled)
        {
            $html_content = $content -replace "`n", "<br>"
            $global:htmlBody += "<font>$html_content</font><br>"
        }
        if($log_disabled -and $console_disabled)
        {
            
            
        }
        elseif($log_disabled)
        {
            Write-Host "$content" -f $color
        }
        else
        {
            if(!($console_disabled -eq 1))
            {            
                $content >> $global:txt_logsfile
                $global:logcontent += "`n$content"
            }
            else 
            {  
                Write-Host "$content" -f $color   
                $content >> $global:txt_logsfile
                $global:logcontent += "`n$content"
            }
        }  
        
    }
    catch
    {
        #Write-Host "Error occurred in LogActivity: $_"
    }
}

function TriggerMail
{
    # Triggers the mail out to the DBA and the team
    param($htmlBody,$CR, $name)
    $smtpServer = "smtp.na.jnj.com"
    $from = "SA-NCSUS-SQLHELPDESK@its.jnj.com"
    $to = "$($name)@its.jnj.com"
    $cc = "PKathirv@its.jnj.com", "DL-ITSUS-GSSQL@ITS.JNJ.COM"
    $subject = "DBRefresh Status - $CR"
    $body = $htmlBody
    $attachment = $pdffilepath

    Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Cc $cc -Subject $subject -Body $body -Attachments $attachment -BodyAsHtml -UseSsl 

}
function GetRITMDetails
{
    # Connects to the IRIS through API call and gets the RITM details
	param
	(
		[string] $RITM,
		[string] $Debug
	)

	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	$Type = "application/json"
	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$headers.Add('Content-Type','application/x-www-form-urlencoded')

	# Specify endpoint uri
	$uri = "https://login.microsoftonline.com/its.jnj.com/oauth2/token?api-version=1.0"

	# Specify HTTP method
	$method = "post"
	
	$BaseURL = "https://jnj-internal-production.apigee.net/apg-001-servicenow/v1/now"
	$IRIS_APIs = Invoke-SQLCmd -query "select Misc from [AdventureWorks2016].[dbo].[Misc] where id = 1" -Server ITSUSRAWSP10439 
	if($IRIS_APIs)
	{
		$bodyJson = "grant_type=client_credentials&client_id=edaff30a-eb13-441a-a9ca-830bc31c165b&client_secret=$($IRIS_APIs)&resource=https%3A//ITS-APP-ISM-IRIS-Prod.jnj.com"
	}
	else
	{
		exit
	}
	

	# Send HTTP request
	$response = Invoke-WebRequest -Headers $headers -Method $method -Uri $uri -Body $bodyJson | ConvertFrom-Json | Select-Object access_token, token_type

	$SNOWSessionHeader = @{'Authorization' = "$($response.token_type) $($response.access_token)"}
    try
    {
	    $sc_req_itemURL = "$($BaseURL)/table/sc_req_item/$($RITM)/variables?nullvariables=true"
	    $RITMJSON = Invoke-RestMethod -Method GET -Uri $sc_req_itemURL -TimeoutSec 100 -Headers $SNOWSessionHeader -ContentType $Type
	    $RITM_RESULTS = $RITMJSON.result
    }
    catch
    {
        Write-Host "`nNo RITM found in IRIS. Please re-verify. Exiting...`n" -f Red
        Exit 1
    }

    $result = [PSCustomObject]@{
        SourceServer = $RITM_RESULTS.$RITM.affected_ci_1
        TargetServer = $RITM_RESULTS.$RITM.affected_ci_9
        TargetEnv = $RITM_RESULTS.$RITM.target_env
        SourceEnv = $RITM_RESULTS.$RITM.sdlc_env
        Info = $RITM_RESULTS.$RITM.DB_refresh_data_file
    }
	return $result
}
function GetCRDetails
{
    # Connects to the IRIS through API call and collects the mentioned CR details
    param
    (
        [string] $CR,
        [string] $Debug
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $Type = "application/json"
    $method = "post"
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add('Content-Type','application/x-www-form-urlencoded')

    # Specify endpoint uri
    $uri = "https://login.microsoftonline.com/its.jnj.com/oauth2/token?api-version=1.0"

    # Specify HTTP method
    $BaseURL = "https://jnj-internal-production.apigee.net/apg-001-servicenow/v1/now"
    if($IRIS_APIs)
    {
    	$bodyJson = "grant_type=client_credentials&client_id=edaff30a-eb13-441a-a9ca-830bc31c165b&client_secret=$($IRIS_APIs)&resource=https%3A//ITS-APP-ISM-IRIS-Prod.jnj.com"
    }
    else
    {
    	exit
    }


    # Send HTTP request
    $response = Invoke-WebRequest -Headers $headers -Method $method -Uri $uri -Body $bodyJson |ConvertFrom-Json|select access_token, token_type

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
        $CHGJSON = Invoke-RestMethod -Method GET -Uri $CHGURL -TimeoutSec 100 -Headers $SNOWSessionHeader -ContentType $Type  
        $CHG_RESULTS = $CHGJSON.result
    }
    catch
    {
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
function TakeSnaps([System.Drawing.Rectangle]$bounds, $path)
{
   # Handles the process of taking screenshots along with the taskbar covering the time details 
   $bmp = New-Object Drawing.Bitmap $bounds.width, $bounds.height
   $graphics = [System.Drawing.Graphics]::FromImage($bmp)
   
   $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)
   
   $bmp.Save($path)
   
   $graphics.Dispose()
   $bmp.Dispose()
}

function ConvertTo-PDF
{
    # Converts the logfile from docx to a pdf format
    param($docFilePath)
    $global:pdffilepath = $docFilePath -replace ".docx", ".pdf"
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false  
    $document = $word.Documents.Open($docFilePath)
    $range = $document.Content
    $range.InsertAfter([System.Environment]::NewLine)  # Ensure new content starts on a new line
    $range.InsertAfter($logcontent)
    $document.Save()
    $document.SaveAs([ref]$pdffilepath, [ref] 17)
    $document.Close()
    
    $word.Quit()

    # Release the COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null


    # Start an instance of Word application
    # $word = New-Object -ComObject Word.Application
    # $word.Visible = $false  # Run Word in the background

    # # Open the Word document
    # $document = $word.Documents.Open($docFilePath)

    # # Convert the document to PDF
    # # 17 corresponds to the PDF format

    # # Close the document and Word application
    # $document.Close()
    # $word.Quit()

    # # Release the COM objects to free up memory
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null


    # $Output = get-content $txtfilepath
    # $pdffilepath = $txtfilepath -replace ".txt", ".pdf"
    # $Word = New-Object -ComObject Word.Application 
    # $Word.Visible = $True 
    # $Doc = $Word.Documents.add() 
    # $Word.Selection.TypeText($Output) 
    # $Doc.SaveAs([ref] $pdffilepath, [ref] 17) 
    # $Word.Close 
    # $Word.Quit
}

function ConvertImagesTo-PDF
{
    # Collocate the captured screensnaps and put them into a pdf file
    param($folder, $CR)
    $docxPath = "$folder\$CR.docx"   
    if(!(test-path -path $docxPath))
    {
        new-item -path $docxPath -ItemType file | Out-Null
    }
    $imagePaths = Get-ChildItem $folder | Where-Object {$_.extension -like "*.png"}  
    $global:pdfPath = $docxPath -replace ".docx", ".pdf"      
    
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false 
    $doc = $word.Documents.Open($docxPath)
    
    foreach ($imagePath in $imagePaths) 
    {
        $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $range.InlineShapes.AddPicture($imagePath.FullName) | out-null
    }
    $doc.SaveAs([ref] $pdfPath, 17) 
    $doc.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
}

function ValidateDetails
{
    # Validate the details collected from RITM and CR
    param([string]$server, $details, [int]$disabled=0)
    if($disabled -eq 1)
    {

    }
    else 
    {   
        LogActivity "`nValidating the Affected CI of RITM and CR..." -htmlenabled 0
        if(!($server -eq $details.AffectedCIs))
        {
            LogActivity "`nThe target server mentioned in RITM($server) is not matching the affected_CI mentioned in the CR$(($details.AffectedCIS)). Exiting" "Red" -htmlenabled 0
            #$htmlBody += "<font>The target server mentioned in RITM($server) is not matching the affected_CI mentioned in the CR$(($details.AffectedCIS))</font>"
            Exit 0
        }
        LogActivity "`nMatching Affected CIs between RITM and CR ? YES" -htmlenabled 0
        if(!($details.state -eq "Implement"))
        {
            LogActivity "`nCR is not in implement state. Exiting" "Red"
            #$htmlBody += "<font>The CR is not in implement state.</font>"
            Exit 0
        }
        LogActivity "CR is in implement state ? YES" 

        $currentDate = Get-Date 
	    if(!($details.StartDate -le $currentDate -and $details.EndDate -ge $currentDate))
        {
            LogActivity "Execution time is not falls within the schedule window of the CR. Exiting" "Red"
            #$htmlBody += "<font>The current time does not fall in between the schedule window of the CR.</font>"
            Exit 0
        }
        LogActivity "Execution time falls within the schedule window ? YES" 
        LogActivity "Validation is completed successfully" "Green"
    }

}
function ServerValidation
{
    # Validates the source and target server accessibility and AG configuration of them
    param([string]$source_server, [string]$target_server)
    LogActivity "`nValidating the servers..."
    foreach($server in @($source_server,$target_server))
    {
        $namedinst = $server.split("\")[0]
        $ping_result = ping $namedinst
        if($ping_result.contains("Request timed out.") -or ($ping_result.contains("could not find host")))
        {
            LogActivity "The server $server is not reachable. Exiting..." "Red"
            #$htmlBody += "<font>The server $server is not reachable.</font>"
            Exit 1
        }

        if((Invoke-Sqlcmd -serverinstance $server -Database "master" -Query "SELECT SERVERPROPERTY('IsHadrEnabled') AS IsHadrEnabled").IsHadrEnabled -eq 1)
        {
            LogActivity "The server $server is an 'Always On' enabled server. Exiting..." "Red"
            #$htmlBody += "<font>The server $server is an AG server</font>"
            Exit 1
        }
    }
    LogActivity "Source and Target servers are reachable and no 'Always ON' configuration detected." "Green"
}
function VersionCompatabilityCheck
{
    # Validates the sql version compatability between the source and target server
    param($sourceserver, $targetserver)
    LogActivity "`nValidating the version compatibility between the source and target servers..."
    $source_majorversion = ((get-sqlagent -serverinstance $sourceserver).serverversion).major
    $target_majorversion = ((get-sqlagent -serverinstance $targetserver).serverversion).major
    if($target_majorversion -ge $source_majorversion)
    {
        LogActivity "The servers are compatible for restore" "Green"
    }
    else 
    {
        LogActivity "The servers are not compatible for restore Activity. The target server is at lower version. Exiting... " "Red"
        #$htmlBody += "<font>The servers are not compatible for restore Activity. The target server has patched with lesser version than the target</font>"
        Exit 1
    }
    if($target_majorversion -eq 10)
    {
        LogActivity "The target server is 2008 SQL server. Automation will not be supported. Exiting..." "Red"
        Exit 1
    }
}
function DatabaseValidation
{
    # Validates the source and taregt databases with multiple specifications like, CDC, TDE, Replication and Multiple files
    param([string]$server, [string[]]$dbs, [int]$source_enabled=0)
    
    $smoServer = New-Object Microsoft.SqlServer.Management.Smo.Server $server
    $dbsize_inTotal = 0
    $logical_data_details = @{}    
    $logical_log_details = @{}  
    $source_db_properties = @{}
    $target_db_properties = @{}
    $remove_multiple_datalog_db = @{}
    $properties_array = @("AutoCreateStatisticsEnabled", "Autoshrink", "AutoUpdateStatisticsEnabled", "Collation", "CompatibilityLevel", "IsFullTextEnabled", "RecoveryModel", "UserAccess", "Version", "Owner", "IsDatabaseSnapshot", "IsReadCommittedSnapshotOn", "SnapshotIsolationState")
    foreach($database in $dbs) 
    {
        $db = $smoServer.Databases[$database]
        if($db) 
        {
            if($db.Status -eq [Microsoft.SqlServer.Management.Smo.DatabaseStatus]::Normal) 
            {
                LogActivity "   Database '$database' is online on the server '$server'." -console_disabled 0
                $dbSize = [math]::Round($db.Size / 1024, 2)
                $dbsize_inTotal += $dbSize 
            } 
            else 
            {
                LogActivity  "Database '$database' is not online on the server '$server'. Status: $($db.Status). Exiting..." "Red"
                #$htmlBody += "<font>The database '$database' is not online on the server '$server'. Status: $($db.Status)</font>"
                Exit 0
            }
            
            #cdc_enabled?
            $cdc_query = "SELECT is_cdc_enabled FROM sys.databases WHERE name = '$($db.Name)'"
            if($smoServer.Databases[$database].ExecuteWithResults($cdc_query).Tables[0].Rows[0].is_cdc_enabled -eq 'True')
            {
                Write-Warning "Database $($db.Name) is CDC enabled. Exiting...Please proceed with manual restore"
                #$htmlBody += "<font>The database $($db.Name) is a CDC enabled.</font>"
                Exit 0
            }
            LogActivity "   Database $($db.Name) is not CDC enabled" -console_disabled 0

            #tde_enabled?
            if($db.EncryptionEnabled -eq 'True')
            {
                Write-Warning "Database $($db.Name) is TDE enabled. Exiting...Please proceed with manual restore"
                #$htmlBody += "<font>The database $($db.Name) is a TDE enabled.</font>"
                Exit 0
            }
            LogActivity "   Database $($db.Name) is not TDE enabled" -console_disabled 0

            #replication_enabled
            $replication_query = "select is_published, is_subscribed, is_merge_published, is_distributor from sys.databases where name='$($db.Name)'"                                  
            $replicationstatus = $db.ExecuteWithResults($replication_query).Tables[0].Rows[0]
            if($replicationstatus.is_published -eq 1 -or $replicationstatus.is_subscribed -eq 1 -or $replicationstatus.is_merge_published -eq 1 -or $replicationstatus.is_distributor -eq 1)
            {
                Write-Warning "    Database $($db.Name) is Replication enabled. Exiting...Please proceed with manual restore"
                #$htmlBody += "<font>The database $($db.Name) is a Replication enabled.</font>"
                Exit 0
            }
            LogActivity "   Database $($db.Name) is not Replication enabled"  -console_disabled 0

            #multiple data files
            if($source_enabled -eq 1)
            {             
                if($db.FileGroups.Files.FileName.count -gt 1)
                {
                    $logical_data_details[$db.name] += $db.Filegroups.Files.Name
                    #$logical_log_details[$db.name] += $db.LogFiles.Name
                    $remove_multiple_datalog_db[$dbs.indexof($database)] += @($database)
                }
                if($db.LogFiles.FileName.count -gt 1)
                {
                    $logical_log_details[$db.name] += $db.Logfiles.Name
                    #$logical_data_details[$db.name] += $db.Filegroups.Files.Name
                    $remove_multiple_datalog_db[$dbs.indexof($database)] += @($database)
                }
                $source_db_properties[$database] += $db.Collation
            }
            if($source_enabled -eq 2)
            {
                foreach($property in $properties_array)
                {
                    $target_db_properties[$database] += @{$property = $($db.$property)}
                }
            
            }
        } 
        else 
        {
            LogActivity "The database '$database' does not exist on server '$server'. Exiting..." "Red"
            #$htmlBody += "<font>The database '$database' does not exist on server '$server'.</font>"
            Exit 0
        }

    }  
    
    return @($logical_data_details, $logical_log_details, $target_db_properties, $remove_multiple_datalog_db, $source_db_properties)
}
function SpaceValidation
{
    # Validates the space availiabilty in the target server
    param($sourceserver, $targetserver, [string[]]$source_dbs, [string[]]$target_dbs)
    $source_smoserver = New-Object Microsoft.SQLServer.Management.SMO.Server $sourceserver
    $target_smoserver = New-Object Microsoft.SQLServer.Management.SMO.Server $targetserver
    
    $db_match_dictionary = @{}
    $count = $source_dbs.count
    $index = 0
    while($index -lt $count)
    {
        $db_match_dictionary[$target_dbs[$index]] = $source_dbs[$index]
        $index++
    }
    $db_size_details = @{}
    $db_target_size_details = @{}
    foreach($key in $db_match_dictionary.Keys)
    {
        $targetdb = $key
        $sourcedb = $db_match_dictionary[$key]
        $targetdb_details = $target_smoserver.Databases[$targetdb]
        $sourcedb_details = $source_smoserver.Databases[$sourcedb]
        if($targetdb_details.FileGroups.Files.count -eq $sourcedb_details.FileGroups.Files.count)
        {
            $source_data_logical_name = @($source_smoserver.Databases[$sourcedb].FileGroups.Files.Name)[0]
            $target_data_physical_path = ($target_smoserver.Databases[$targetdb].FileGroups.Files | Where-Object {$_.id -eq 1}).FileName
            Write-Host "`n$source_data_logical_name ------> $target_data_physical_path"
            foreach($file in @($targetdb_details.FileGroups.Files.FileName))
            {
                $index = $targetdb_details.FileGroups.Files.FileName.indexof($file)                
                $drive = (split-path $file -Qualifier).trim(":")
                $size = $sourcedb_details.FileGroups.Files.Size[$index]
                $db_size_details[$drive] += $size
                $target_size = $targetdb_details.FileGroups.Files.Size[$index]
                $db_target_size_details[$drive] += $target_size
            }
        }
        if($targetdb_details.LogFiles.count -eq $sourcedb_details.LogFiles.count)
        {
            $source_log_logical_name = @($source_smoserver.Databases[$sourcedb].LogFiles.Name)[0]
            $target_log_physical_path = ($target_smoserver.Databases[$targetdb].LogFiles | Where-Object {$_.id -eq 2}).FileName
            Write-Host "$source_log_logical_name ------> $target_log_physical_path"

            foreach($file in @($targetdb_details.LogFiles.FileName))
            {
                $index = $targetdb_details.LogFiles.FileName.indexof($file)
                $drive = (split-path $file -Qualifier).trim(":")
                $size = $sourcedb_details.LogFiles.Size[$index]
                $db_size_details[$drive] += $size
                $target_size = $targetdb_details.LogFiles.Size[$index]
                $db_target_size_details[$drive] += $target_size
            }
        }
        if($targetdb_details.FileGroups.Files.count -gt $sourcedb_details.FileGroups.Files.count)
        {
            $source_data_logical_name = @($source_smoserver.Databases[$sourcedb].FileGroups.Files.Name)[0]
            $target_data_physical_path = ($target_smoserver.Databases[$targetdb].FileGroups.Files | Where-Object {$_.id -eq 1}).FileName
            Write-Host "`n$source_data_logical_name ------> $target_data_physical_path"
        
            foreach($file in @($sourcedb_details.FileGroups.Files.FileName))
            {
                $index = $targetdb_details.FileGroups.Files.FileName.indexof($file)    
                $drive = (split-path $(@($targetdb_details.FileGroups.Files.FileName)[$sourcedb_details.FileGroups.Files.FileName.indexof($file)]) -Qualifier).trim(":")
                $size = $sourcedb_details.FileGroups.Files.Size[$index]
                $db_size_details[$drive] += $size
                $target_size = $targetdb_details.FileGroups.Files.Size[$index]
                $db_target_size_details[$drive] += $target_size
            }
        }
        if($targetdb_details.LogFiles.count -gt $sourcedb_details.LogFiles.count)
        {
            $source_log_logical_name = @($source_smoserver.Databases[$sourcedb].LogFiles.Name)[0]
            $target_log_physical_path = ($target_smoserver.Databases[$targetdb].LogFiles | Where-Object {$_.id -eq 2}).FileName
            Write-Host "$source_log_logical_name ------> $target_log_physical_path"
            
            foreach($file in @($sourcedb_details.LogFiles.FileName))
            {
                $index = $targetdb_details.LogFiles.FileName.indexof($file)
                $drive = (split-path $(@($targetdb_details.LogFiles.FileName)[$sourcedb_details.LogFiles.FileName.indexof($file)]) -Qualifier).trim(":")
                $size = $sourcedb_details.LogFiles.Size[$index]
                $db_size_details[$drive] += $size
                $target_size = $targetdb_details.LogFiles.Size[$index]
                $db_target_size_details[$drive] += $target_size
            }
        }
        if($targetdb_details.FileGroups.Files.count -lt $sourcedb_details.FileGroups.Files.count)
        {
            $source_data_logical_name = @($source_smoserver.Databases[$sourcedb].FileGroups.Files.Name)[0]
            $target_data_physical_path = ($target_smoserver.Databases[$targetdb].FileGroups.Files | Where-Object {$_.id -eq 1}).FileName
            Write-Host "`n$source_data_logical_name ------> $target_data_physical_path"
        
            foreach($file in @($sourcedb_details.FileGroups.Files))
            {
                try
                {    
                    $index = $sourcedb_details.FileGroups.Files.indexof($file)
                    $drive = (split-path $(@($targetdb_details.FileGroups.Files.FileName)[$index]) -Qualifier).trim(":")
                    $size = $sourcedb_details.FileGroups.Files.Size[$index]
                    $db_size_details[$drive] += $size
                    $target_size = $targetdb_details.FileGroups.Files.Size[$index]
                    $db_target_size_details[$drive] += $target_size
                    
                }
                catch
                {
                    $target_file = $targetdb_details.FileGroups.Files
                    $drive = (split-path $(@($target_file.FileName)[0]) -Qualifier).trim(":")
                    $size = $sourcedb_details.FileGroups.Files.Size[$index]
                
                    $db_size_details[$drive] += $size
                }
            }
        }
        if($targetdb_details.LogFiles.count -lt $sourcedb_details.LogFiles.count)
        {
            $source_log_logical_name = @($source_smoserver.Databases[$sourcedb].LogFiles.Name)[0]
            $target_log_physical_path = ($target_smoserver.Databases[$targetdb].LogFiles | Where-Object {$_.id -eq 2}).FileName
            Write-Host "$source_log_logical_name ------> $target_log_physical_path"

            foreach($file in @($sourcedb_details.LogFiles.FileName))
            {
                TRY
                {
                    $index = $sourcedb_details.LogFiles.FileName.indexof($file)
                    $drive = (split-path $(@($targetdb_details.LogFiles.FileName)[$index]) -Qualifier).trim(":")
                    $size = $sourcedb_details.LogFiles.Size[$index]
                    $db_size_details[$drive] += $size
                    $target_size = $targetdb_details.LogFiles.Size[$index]
                    $db_target_size_details[$drive] += $target_size

                }
                catch
                {
    
                    $target_file = $targetdb_details.LogFiles
                    $drive = (split-path $(@($target_file.FileName)[0]) -Qualifier).trim(":")
                    $size = $sourcedb_details.LogFiles.Size[$index]
                
                    $db_size_details[$drive] += $size
                }
            }
        }
    }
    $namedinst = $targetserver.split("\")[0]
    $target_volume = Invoke-Command -ComputerName $namedinst -ScriptBlock{
        return Get-Volume
    }
    $free_space = @{}
    foreach($key in $db_size_details.Keys)
    {
        $source_size_GB = ($($db_size_details[$key])*1024)/1GB
        $target_total_size = [Math]::Round(($target_volume | Where-Object {$_.DriveLetter -eq $key}).Size/1GB)
        $target_free_size = [Math]::Round(($target_volume | Where-Object {$_.DriveLetter -eq $key}).SizeRemaining/1GB)
        $target_free_size = $target_free_size + ($($db_target_size_details[$key])*1024)/1GB
        $restore_free_size = $target_free_size - $source_size_GB
        #$free_size_percent = [Math]::Round(($target_total_size/$restore_free_size), 2)
        $free_size_percent = [Math]::Round(($restore_free_size/$target_total_size), 2)
        $expected_free_size = ($target_total_size*10)/100
        if(!($restore_free_size -ge $expected_free_size))
        {
            LogActivity  "`nThe drive $key does not have enough space to accomodate the restore activity" #"Red"
            #$htmlBody += "<font>The drive $key does not have enough space to accomodate the restore activity</font>"
            Exit 0
        }
        $free_space[$key] += @($restore_free_size, $free_size_percent)
    }
    return $free_space
}
function FindNAS
{
    # Finds the nearest NAS location based on the target server location
    param($server)
    
    $namedinst = $server.split("\")[0]
    $region = Invoke-command -ComputerName $namedinst -ScriptBlock{
        return $($env:Region)
    }
    if($region)
    {
        if($region -ieq "na")
        {
            #LogActivity "The Target server located in the NA region."
            return @("NA", "\\itsusrac1ts1\sqldbrefresh_ra_1\REFRESH")
        }
        elseif($region -ieq "eu")
        {
            #LogActivity "The Target server located in the EMEA region."
            return @("EMEA", "\\itsbebec1ts1.jnj.com\sqldbrefresh_be_1\REFRESH")
        }
        elseif($region -ieq "ap")
        {
            #LogActivity "The Target server located in the ASPAC region."
            return @("ASPAC", "\\awssgdwufsxn01.jnj.com\sqldbrefresh_sg_1\REFRESH")
        }
    }
}
function CreateFolders_in_NAS
{
    # Creates the folders in the selected NAS location 
    param(
        [string[]]$dbs, 
        [string]$CR, 
        [string]$NAS,
        [int]$type
    )

    if($type -eq 1)
    {
        $collection = "Source"
    }
    elseif($type -eq 2)
    {
        $collection = "Target"
    }

    $main_workfolder = "$workfolder" + "\" + $($collection) + "_Files"
    New-Item -path $main_workfolder -ItemType Directory | Out-Null
    foreach($db in $dbs)
    {
        $path = "$main_workfolder\$db"
        New-Item -path $path -ItemType Directory | Out-Null
    }
}
function Invoke_SQLDBBackup 
{ 
    # Takes the full backups of the source and target dbs
    param
    (
        [Parameter(Mandatory=$true,Position=0)][String]$SQLServer,
        [Parameter(Mandatory=$true,Position=1)][String]$BackupDirectory, 
        [Parameter(Mandatory=$true,Position=2)][String[]]$dbList
        #[Parameter(Mandatory=$true,Position=3)][System.Management.Automation.Credential()]$credential
    ) 

    try
    {
        $namedinst = $SQLServer.split("\")[0]
        foreach($db in $dbList)
        {
            LogActivity "`nDatabase Name: `'$db`'"
            LogActivity "Start Time: $(get-date -f "MM/dd/yyyy hh:mm:ss") EST"
            $backup_disk_path = "'$BackupDirectory\$($db)\$($db).bak'"

            $time = Measure-Command{
                $backup_query = "BACKUP DATABASE `"$db`"
                TO DISK = $backup_disk_path
                WITH COPY_ONLY, init;"
                try
                {
                    Invoke-SQLCmd -ServerInstance $namedinst -Query $backup_query -ErrorAction Stop
                }
                catch
                {
                    Write-Host "Error Occurred in taking the backup of the db $db : $_"
                    Exit 0
                }
            }
            LogActivity "End Time: $(get-date -f "MM/dd/yyyy hh:mm:ss") EST"
            LogActivity "File Path: $BackupDirectory\$db\$db.bak"
            #start-sleep -s 45
            LogActivity "Validating the backup..."
            $verify_query = "restore verifyonly from disk = '$BackupDirectory\$db\$db.bak'"
            try 
            {
                Invoke-Sqlcmd -ServerInstance $namedinst -Query $verify_query -ErrorAction Stop
                LogActivity "Backup Validation output: The backup set on file 1 is valid" "Green"
            }
            catch 
            {
                LogActivity "Backup Validation output : Verification failed. Exiting" "Red"
                Exit 0
            }
            #$fileonly_query = "restore filelistonly from disk = '$BackupDirectory\$db\$db.bak'"
            #$result = Invoke-Sqlcmd -serverinstance $namedinst -query $fileonly_query | select-object LogicalName, PhysicalName
            #LogActivity "`n   Time taken for taking the backup $db : $([Math]::Round($time.totalMinutes,5)) minutes"
            #LogActivity "`n   Files in the db $db : "
            #foreach($file_value in $result)
            #{
            #    LogActivity "          $($file_value.LogicalName)   *** $($file_value.PhysicalName)"
            #}
            #
            #Logactivity "`n"
        }        

    }
    catch
    {
        LogActivity "Error Occurred : $_ "
        #$htmlBody += "<font>Error Occurred : $_ </font>"
        Exit 0
    }
     
}
function Copy_Logins
{
    # Copies out the users, roles and permissions of the target dbs to the NAS location
    param($targetserver, $backuplocation, $targetdbs)

    $SqlConnection = New-Object System.Data.SQLClient.SQLConnection
    $SqlCommand = New-Object System.Data.SQLClient.SqlCommand;

    $query_db_users = "/*SCRIPT OUT DATABASE USERS*/
                        SELECT 'IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '''+ DP.NAME collate database_default +''') CREATE USER ['+ DP.NAME collate database_default +'] FOR LOGIN [' + SP.NAME+']'
                        AS '/*DB USERS*/'
                        FROM SYS.DATABASE_PRINCIPALS DP
                        JOIN SYS.SERVER_PRINCIPALS SP
                        ON DP.SID =SP.SID AND DP.PRINCIPAL_ID > 4"

    $query_db_role = "SELECT 'ALTER ROLE ['+USER_NAME(RM.ROLE_PRINCIPAL_ID) +'] ADD MEMBER [' + USER_NAME(RM.MEMBER_PRINCIPAL_ID) +']'
                        AS '/*DB ROLE MEMBERS*/'
                        FROM SYS.DATABASE_ROLE_MEMBERS RM
                        JOIN SYS.DATABASE_PRINCIPALS DP
                        ON RM.MEMBER_PRINCIPAL_ID =DP.PRINCIPAL_ID AND RM.MEMBER_PRINCIPAL_ID > 4
                        JOIN SYS.SERVER_PRINCIPALS SP
                        ON DP.SID =SP.SID"

    $query_db_perms = "/* SCRIPT DB LEVEL PERMISSIONS */
                        SELECT 
                        STATE_DESC+' '+ DM.PERMISSION_NAME+ ' TO ['+USER_NAME(DM.GRANTEE_PRINCIPAL_ID)+']'+
                        CASE DM.STATE
                        	WHEN 'W' THEN ' WITH GRANT OPTION'
                        	ELSE ''
                        END
                        AS '/*DB LEVEL PERMISSIONS*/'
                        FROM SYS.DATABASE_PERMISSIONS DM
                        JOIN SYS.DATABASE_PRINCIPALS DP
                        ON DM.GRANTEE_PRINCIPAL_ID =DP.PRINCIPAL_ID AND DM.GRANTEE_PRINCIPAL_ID >4 AND DM.CLASS=0
                        JOIN SYS.SERVER_PRINCIPALS SP
                        ON DP.SID=SP.SID"

    $query_db_objperms = "/* SCRIPT DB OBJECT LEVEL PERMISSIONS */
                            SELECT 
                            CASE DM.STATE
                            	WHEN 'W' THEN 'GRANT'
                            	ELSE DM.STATE_DESC
                            END
                            +' '+ 
                            CASE DM.PERMISSION_NAME
                            	WHEN 'REFERENCES' THEN CASE DM.MINOR_ID 
                            								WHEN 0 THEN DM.PERMISSION_NAME
                            								ELSE DM.PERMISSION_NAME+'('+COL_NAME(DM.MAJOR_ID,DM.MINOR_ID)+')'
                            							END
                            	ELSE DM.PERMISSION_NAME 
                            END
                            +' ON OBJECT::[' + OBJECT_SCHEMA_NAME(DM.MAJOR_ID)+'].['+OBJECT_NAME(DM.MAJOR_ID)+']'+
                            ' TO ['+USER_NAME(DM.GRANTEE_PRINCIPAL_ID)+']'+
                            CASE DM.STATE
                            	WHEN 'W' THEN ' WITH GRANT OPTION'
                            	ELSE ''
                            END
                            AS '/*DB OBJECT LEVEL PERMISSIONS*/'
                            FROM SYS.DATABASE_PERMISSIONS DM
                            JOIN SYS.DATABASE_PRINCIPALS DP
                            ON DM.GRANTEE_PRINCIPAL_ID =DP.PRINCIPAL_ID AND DM.GRANTEE_PRINCIPAL_ID >4 AND DM.CLASS=1
                            JOIN SYS.SERVER_PRINCIPALS SP
                            ON DP.SID=SP.SID"

    $query_db_schemaperms = "/* SCRIPT DB SCHEMA LEVEL PERMISSIONS */
                                SELECT 
                                CASE DM.STATE
                                	WHEN 'W' THEN 'GRANT '
                                	ELSE DM.STATE_DESC
                                END
                                +' '+ 
                                DM.PERMISSION_NAME
                                +' ON SCHEMA::[' + SCHEMA_NAME(DM.MAJOR_ID)+']'+
                                ' TO ['+USER_NAME(DM.GRANTEE_PRINCIPAL_ID)+']'+
                                CASE DM.STATE
                                	WHEN 'W' THEN ' WITH GRANT OPTION'
                                	ELSE ''
                                END
                                AS '/*DB SCHEMA LEVEL PERMISSIONS*/'
                                FROM SYS.DATABASE_PERMISSIONS DM
                                JOIN SYS.DATABASE_PRINCIPALS DP
                                ON DM.GRANTEE_PRINCIPAL_ID =DP.PRINCIPAL_ID AND DM.GRANTEE_PRINCIPAL_ID >4 AND DM.CLASS=3
                                JOIN SYS.SERVER_PRINCIPALS SP
                                ON DP.SID=SP.SID"

    $query_other_perms = "/*ANY OTHER PERMISSIONS*/
                            SELECT 
                            CASE DM.STATE
                            	WHEN 'W' THEN 'GRANT'
                            	ELSE DM.STATE_DESC
                            END
                            +' '+ 
                            DM.PERMISSION_NAME 
                            +' '+
                            CASE DM.CLASS
                            	WHEN 4 THEN 'ON ' + (SELECT RIGHT(TYPE_DESC, 4) + '::[' + NAME FROM SYS.DATABASE_PRINCIPALS WHERE PRINCIPAL_ID = DM.MAJOR_ID) + '] '
                            	WHEN 5 THEN 'ON ASSEMBLY::[' + (SELECT NAME FROM SYS.ASSEMBLIES WHERE ASSEMBLY_ID = DM.MAJOR_ID) + '] '
                            	WHEN 6 THEN 'ON TYPE::[' + (SELECT NAME FROM SYS.TYPES WHERE USER_TYPE_ID = DM.MAJOR_ID) + '] '
                                WHEN 10 THEN 'ON XML SCHEMA COLLECTION::[' + (SELECT SCHEMA_NAME(SCHEMA_ID) + '.' + NAME FROM SYS.XML_SCHEMA_COLLECTIONS WHERE XML_COLLECTION_ID = DM.MAJOR_ID) + '] '
                            	WHEN 15 THEN 'ON MESSAGE TYPE::[' + (SELECT NAME FROM SYS.SERVICE_MESSAGE_TYPES WHERE MESSAGE_TYPE_ID = DM.MAJOR_ID) + '] '
                                WHEN 16 THEN 'ON CONTRACT::[' + (SELECT NAME FROM SYS.SERVICE_CONTRACTS WHERE SERVICE_CONTRACT_ID = DM.MAJOR_ID) + '] '
                                WHEN 17 THEN 'ON SERVICE::[' + (SELECT NAME FROM SYS.SERVICES WHERE SERVICE_ID = DM.MAJOR_ID) + '] '
                                WHEN 18 THEN 'ON REMOTE SERVICE BINDING::[' + (SELECT NAME FROM SYS.REMOTE_SERVICE_BINDINGS WHERE REMOTE_SERVICE_BINDING_ID = DM.MAJOR_ID) + '] '
                                WHEN 19 THEN 'ON ROUTE::[' + (SELECT NAME FROM SYS.ROUTES WHERE ROUTE_ID = DM.MAJOR_ID) + '] '
                                WHEN 23 THEN 'ON FULLTEXT CATALOG::[' + (SELECT NAME FROM SYS.FULLTEXT_CATALOGS WHERE FULLTEXT_CATALOG_ID = DM.MAJOR_ID) + '] '
                                WHEN 24 THEN 'ON SYMMETRIC KEY::[' + (SELECT NAME FROM SYS.SYMMETRIC_KEYS WHERE SYMMETRIC_KEY_ID = DM.MAJOR_ID) + '] '
                                WHEN 25 THEN 'ON CERTIFICATE::[' + (SELECT NAME FROM SYS.CERTIFICATES WHERE CERTIFICATE_ID = DM.MAJOR_ID) + '] '
                                WHEN 26 THEN 'ON ASYMMETRIC KEY::[' + (SELECT NAME FROM SYS.ASYMMETRIC_KEYS WHERE ASYMMETRIC_KEY_ID = DM.MAJOR_ID) + ']'
                            END COLLATE DATABASE_DEFAULT
                            +
                            ' TO ['+USER_NAME(DM.GRANTEE_PRINCIPAL_ID)+']'+
                            CASE DM.STATE
                            	WHEN 'W' THEN ' WITH GRANT OPTION'
                            	ELSE ''
                            END
                            AS '/*OTHER PERMISSIONS*/'
                            FROM SYS.DATABASE_PERMISSIONS DM
                            JOIN SYS.DATABASE_PRINCIPALS DP
                            ON DM.GRANTEE_PRINCIPAL_ID =DP.PRINCIPAL_ID AND DM.GRANTEE_PRINCIPAL_ID >4 AND DM.CLASS >=4
                            JOIN SYS.SERVER_PRINCIPALS SP
                            ON DP.SID=SP.SID"
        
    $queries_array = @($query_db_users, $query_db_role, $query_db_perms, $query_db_objperms, $query_db_schemaperms, $query_other_perms)
    try 
    {
        foreach($db in $targetdbs)
        {
            $outfile = $backuplocation + "\" + $db + "\" + $db + "_login.sql"
            if(!(Test-Path -path $outfile))
            {
                New-Item -path $outfile -ItemType file | out-null
            }
            $connectionString = "Server=$targetserver;Database=$db;Integrated Security=True"
            $SqlConnection.ConnectionString = $connectionString
            $SqlCommand.Connection = $SqlConnection
            foreach($query in $queries_array)
            {
                $SqlCommand.CommandText = $query
                $SqlConnection.Open()
                $table = $SqlCommand.ExecuteReader()
                While ($table.Read())
                {
                    $table[0]+' '+"`r`n"+"GO"| Out-File -Append -FilePath $outfile
                }
                $SqlConnection.Close();
            }
            LogActivity "File Path: $outfile" -console_disabled 0
        }
    }
    catch 
    {
        LogActivity "`nError occurred in retriving the logins and permissions out for the db $db" "Red" 
        #$htmlBody += "<font>Error occurred in retriving the logins and permissions out for the db $db</font>"
        Exit 0
    }

}
function get_data_log_filepath
{
    # Collects the details of the physicalnames of the target database
    param($db)
    $filepathsql_query = "SELECT mf.physical_name AS FilePath FROM sys.master_files mf 
                    INNER JOIN sys.databases db ON mf.database_id = db.database_id 
                    WHERE mf.type IN (0, 1) and db.name like '$db'"
    
    $result_filepath = Invoke-Sqlcmd -ServerInstance $targetserver -Query $filepathsql_query
    $extensionOrder = @{
            '.mdf' = 1
            '.ndf' = 2
            '.ldf' = 3
            }


    $sorted_result_filepath = $result_filepath.FilePath | Sort-Object { $extensionOrder[[System.IO.Path]::GetExtension($_)] }

    return $sorted_result_filepath
}
function get_logicalname
{
    # Collects the details of the logical names of the source databases
    param($sourcedb, $backuplocation)
    $logicalname_query = "RESTORE FILELISTONLY
                            FROM DISK = N'$($backuplocation)\$($sourcedb).bak'"
    $result_logicalname = Invoke-Sqlcmd -ServerInstance $targetserver -query $logicalname_query
    return $result_logicalname

}
function ConstructQuery
{
    # Constructs the SQL Query to execute the restore
    param($targetdb, $sourcedb, $filepath, $log_filepaths, $logicalname, $source_backupdirectory)
    $backuplocation = $source_backupdirectory
    $timestamp = Get-date -format "MMddyyyy"
    $data_terminal = "$($targetdb)_automation$timestamp.$($filepath[0].split("\")[-1].split(".")[1])"
    $datafilepath = join-path (split-path $filepath[0] -parent)  $data_terminal
    $log_terminal = "$($targetdb)_log_automation$timestamp.$($log_filepaths[0].split("\")[-1].split(".")[1])"
    $logfilepath = join-path (split-path $log_filepaths[0] -parent)  $log_terminal
    $datalogicalname = $logicalname[0].LogicalName
    $loglogicalname = $logicalname[1].LogicalName
    $diskpath = "$backuplocation"

    $query_to_restore = "USE [master]
                        Alter database [$targetdb] set Restricted_user with rollback immediate 
                        RESTORE DATABASE [$targetdb] FROM  DISK = N'$diskpath' WITH FILE = 1,  
                        MOVE N'$datalogicalname' TO N'$datafilepath', 
                        MOVE N'$loglogicalname' TO N'$logfilepath', 
                        NOUNLOAD,  REPLACE,  STATS = 5
                        GO
                        Alter database [$targetdb] set multi_user"
    return $query_to_restore
}

try
{
    #---------------------------------------------------loading the assemblies--------------------------------------------------------#
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null 
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoExtended") | Out-Null 
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null 
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoEnum") | Out-Null

    #$password = Read-Host -Prompt "Enter your password " -AsSecureString
    $username = $($env:username)
    #$global:Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $global:htmlBody = ""

    $delimiter = "*******************************************************************************************************************************"
    Write-Host "`n$delimiter"
    Write-Host "DB Refresh: Input Parameters`n"
    Write-Host "DESCRIPTION
    This script takes four inputs from the user :  RITM number, CR number, Source DB(s), Target DB(s)
    Validates the correctness of the entered RITM and CR numbers. Dynamically picks the value of source and target servers from the RITM.
    Validates the servers accessibility and compatibility.
    Validates the source and target DB(s) with list of specifications.
    Check the space availability in the target server.
    Finds the nearest NAS location and saves the full backups of both source and target DB(s). 
    Saves the login scripts and properties of target DB(s) in the NAS location.
    Executes the restore of source DB(s) to target DB(s) respectively. Post retsore, logins & permissions will be re-applied.
NOTES
    Ensure that the script is run with administrative privileges.
    This version of script is not handling the restore of DB(s) with multiple data/log files`n" -f CYan
    #---------------------------------------------------RITM details--------------------------------------------------------#
    #$RITM = Read-Host "Enter the RITM Number " #"RITM000022712633" "RITM000022529841" 
    #$RITM = $RITM.trim()
    #$RITMDetails = GetRITMDetails -RITM $RITM
    #if(!($RITMDetails.SourceServer))
    #{
    #    Write-Host "`nNo RITM found in IRIS. Please re-verify. Exiting..." -f Red
    #    $htmlBody += "<font>No Data extracted from the RITM API call. </font>"
    #    Exit 0
    #}
    #else
    #{
    #    Write-Host "API Call is successful and RITM validated successfully " -f Green
    #    Write-Host "Source Server: $($RITMDetails.SourceServer) ($($RITMDetails.SourceEnv))"
    #    Write-Host "Target Server: $($RITMDetails.TargetServer) ($($RITMDetails.TargetEnv))"
    #    $target_env = $RITMDetails.TargetEnv
    #    if($target_env -ieq "prod" -or $target_env -ilike "*prod*")
    #    { 
    #       Write-Host "The target server is a production environment. Exiting..." -f Red
    #       Exit 0
    #    }
    #}
    
    #---------------------------------------------------CR details--------------------------------------------------------#
    #$CRNumber = Read-Host "`nEnter the CR Number " #"CHG000011177902" "CHG000011152665"
    #$CRNumber = $CRNumber.trim()
    #$CRDetails = GetCRDetails -CR $CRNumber 
    #if(($CRDetails[1] -eq $null))
    #{
    #    LogActivity  "Exiting..." "Red"
    #    $htmlBody += "<font>No data fetched from CR API call.</font>"
    #    Exit 0
    #}
    #else
    #{
    #    Write-Host "CR number is validated." -f Green
    #}

    #---------------------------------------------------Get the input details--------------------------------------------------------#
    #$sourceserver = $RITMDetails.SourceServer #Read-Host "`nEnter the source server "
    #$targetserver = $RITMDetails.TargetServer #Read-Host "`nEnter the target server "
    $sourceserver = Read-Host "`nEnter the source server "
    $targetserver = Read-Host "`nEnter the target server "
    #Write-Host "`nValidating the details collected from RITM and CR..."
    #ValidateDetails $targetserver $CRDetails -disabled 1

    Write-Host "`nInput Source and Target Database Name(s)."
    #Start-Sleep -s 2
    Write-Host "For multiple dbs (max 3), enter db names with comma separated in the exact same order. (Example: Source: DB1, DB2 Target: DB1, DB2)"
    #Start-Sleep -s 2
    $source_dbs = @(Read-Host "`nEnter the source dbs") #"#@("Refresh_Test")
    $target_dbs = @(Read-Host "Enter the target dbs") #"#@("TestDB1")
    #$source_dbs = @("Refresh_Test")
    #$target_dbs = @("TestDB1")
    $source_dbs = @($source_dbs -split "," | ForEach-Object{$_.trim()})
    $target_dbs = @($target_dbs -split "," | ForEach-Object{$_.trim()})
    if(!($source_dbs.count -eq $target_dbs.count))
    {
        Write-Host "The number of source and target dbs are not matching..Exiting.." "Red"
        Exit 1
    }
    if($source_dbs.count -gt 3)
    {
        Write-Host "Please provide only 3 databases maximum in a time. Exiting..." "Red"
        Exit 1
    }

    Write-Host "`nReview Input Parameters`n"
    Write-Host "$delimiter"
    Write-Host  "RITM : $RITM"  -f Cyan
    Write-Host  "CR : $CRNumber" -f Cyan
    Write-Host  "Source Server : $sourceserver" -f Cyan
    Write-Host  "Target Server : $targetserver" -f Cyan
    Write-Host  "Source DB(s) : $($($source_dbs | foreach-object {"`'$_`'"}) -join ', ')" -f Cyan
    Write-Host  "Target DB(s) : $($($target_dbs | foreach-object {"`'$_`'"}) -join ', ')" -f Cyan
    Write-Host "$delimiter`n"

    #Start-Sleep -s 2
    #$proceed_prompt = Read-Host "Press Y or y to proceed. N or n to exit"
    #if(!($proceed_prompt -ieq "y"))
    #{
    #    Write-Host "`nExiting...`n"
    #    Exit 0
    #}
    #Write-Host "`nproceeding..."
    #Start-Sleep -s 2
    #-----------------------------------------------------------------------------------------------------------#
    #$targetenv = $RITMDetails.TargetEnv
    #$db_info = $RITMDetails.Info -split "`n"
    #$sourceserver = "OPCITSNAW00R1"
    #$source_dbs = @($db_info[0] -split ",")
    #$target_dbs = @($db_info[1] -split ",")
     
    #----------------------------------------------------Create the folders and files-------------------------------------------------------#
    $timeStamp = Get-Date -format "yyyy-MM-dd_HH_mm_ss"

    $local_basefolder = "D:\DBRefresh\Logs"
    if(!(Test-Path -path $local_basefolder))
    {
        New-Item -path $local_basefolder -ItemType Directory | Out-Null
    }

    $local_workfolder = "$local_basefolder\$($CRNumber)_$($timeStamp)"
    if(!(Test-Path -path $local_workfolder))
    {
        New-Item -path $local_workfolder -ItemType Directory | Out-Null
    }

    $global:logsfile = "$local_workfolder\DBRefresh_$($CRNumber).docx"
    $global:txt_logsfile = "$local_workfolder\DBRefresh_$($CRNumber).txt"
    if(!(Test-Path -path $logsfile))
    {
        New-Item -path $logsfile -ItemType File | Out-Null
    }
        if(!(Test-Path -path $txt_logsfile))
    {
        New-Item -path $txt_logsfile -ItemType File | Out-Null
    }

    #-----------------------------------------------------Print the details------------------------------------------------------#
    $global:logcontent = ""
    LogActivity  "`n$delimiter`nDB Refresh started on : $(Get-Date -f 'MM/dd/yyyy HH:mm:ss') EST`nExecuted By: $username`n$delimiter" 

    #------------------------------------------------------PreValidations----------------------------------------------- 
    LogActivity "`nReview Input Parameters`n" -console_disabled 0
    LogActivity "$delimiter" -console_disabled 0
    LogActivity  "RITM : $RITM" -console_disabled 0
    LogActivity  "CR : $CRNumber" -console_disabled 0
    LogActivity  "Source Server : $sourceserver" -console_disabled 0
    LogActivity  "Target Server : $targetserver" -console_disabled 0
    LogActivity  "Source DB(s) : $($($source_dbs | foreach-object {"`'$_`'"}) -join ', ')" -console_disabled 0
    LogActivity  "Target DB(s) : $($($target_dbs | foreach-object {"`'$_`'"}) -join ', ')" -console_disabled 0
    LogActivity "$delimiter`n"  -console_disabled 0

    ServerValidation $sourceserver $targetserver
    Start-Sleep -s 1
    VersionCompatabilityCheck $sourceserver $targetserver
    Start-Sleep -s 1

    LogActivity "`nValidating the source and target DB special handling parameters..."
    $multiple_file_details = DatabaseValidation $sourceserver $source_dbs 1
    $remove_dbs = $multiple_file_details[3]
    $remove_dbs_sorted = $remove_dbs.getenumerator() | foreach-object {
        $_.Value = $_.Value | sort-object | Get-Unique
        $_}

    $remove_dbs_sorted | foreach-object {
        LogActivity "`n    DB '$($_.Value)' has multiple data/log files. Exiting and Proceed with Manual Restore ..." "Red"
        Exit 0
    }
    
    $source_dbs = @($source_dbs | where-object {@($($remove_dbs_sorted.Value)) -notcontains $_})
    $target_dbs = @($target_dbs | foreach-object { $index = [array]::IndexOf($target_dbs, $_)
        if (-not $remove_dbs.ContainsKey($index)) {
            $_
        }})

    $source_db_properties = $multiple_file_details[4]
    $target_db_properties = DatabaseValidation $targetserver $target_dbs 2
    if($source_dbs.count -eq 0 -or $target_dbs.count -eq 0)
    {
        LogActivity "Exiting as no db to continue..." "Red"
        Exit 0

    }
    LogActivity "Validated successfully. Database(s) are online. No TDE, CDC, Replication enabled on Source and Target." "Green"
    Start-Sleep -s 1

    LogActivity "`nValidating the space availability on the target server..."
    $space_details = SpaceValidation $sourceserver $targetserver $source_dbs $target_dbs
    LogActivity "`nSpace validated successfully." "Green"
    Start-Sleep -s 1

    LogActivity "`n$delimiter`nPre-validation succeeded ::: Ready to proceed with implementation.`n$delimiter`n" "Green"
    #$cr_prompt = Read-Host "Move the CR to implement state and press Enter once done "
    #if($cr_prompt -eq "")
    #{
    #    $secondary_cr_details = GetCRDetails -CR $CRNumber
    #    ValidateDetails $targetserver $secondary_cr_details 
    #}
    #else 
    #{
    #    LogActivity "`nExiting because of wrong input" "Red"
    #    Exit 0
    #}
    
    #---------------------------------------------------------Take the prevalidation snap--------------------------------------------
    #Start-Sleep -s 2
    Add-Type -AssemblyName System.Drawing
    $FilePath = "$local_workfolder\Image1_PreImplementation.png"
    $bounds = [System.Drawing.Rectangle]::FromLTRB(0, 0, 2000, 2000)
    #TakeSnaps $bounds $FilePath
    
    #-------------------------------------------------------------Start of the Implememtation----------------------------------------------------------------------------
    #*****************find the NAS location******************
    LogActivity "`n***** Locating the nearest NAS File share and directory creation *****"
    $nas_location = (FindNAS $targetserver)[1]
    LogActivity "Closest NAS location: $((FindNAS $targetserver)[0])"
    #Start-Sleep -s 1

    $basefolder = "$($CRNumber)_$($timeStamp)"
    $workfolder = "$nas_location\$basefolder"
    New-Item -path $workfolder -ItemType Directory | Out-Null
    $source_backupdirectory = "$workfolder\Source_Files"
    $target_backupdirectory = "$workfolder\Target_Files"

    #*****************create folders in the NAS location******************
    LogActivity "`nCreating directories in the NAS Location to place the backup(s)..."
    LogActivity "Source Directory: $source_backupdirectory"
    LogActivity "Target Directory: $target_backupdirectory"
    CreateFolders_in_NAS -dbs $source_dbs -CR $CRNumber -NAS $workfolder -type 1
    CreateFolders_in_NAS -dbs $target_dbs -CR $CRNumber -NAS $workfolder -type 2
    LogActivity "`nDirectories created in the NAS Location successfully." "Green" -console_disabled 0
    #LogActivity "`n$delimiter`n"
    #Start-Sleep -s 2

    #*****************take the backups and store in the NAS location******************
    try
    {
        LogActivity "`n$delimiter`nSource Backup(s)`n$delimiter`n"
        LogActivity "Initiating Source DB Backup(s) ..."
        Invoke_SQLDBBackup $sourceserver $source_backupdirectory $source_dbs 
        LogActivity "`nSource backup completed successfully." "Green"
        #start-sleep -s 45
        #LogActivity "   Source Database Backups are taken successfully to the location $source_backupdirectory" "CYAN"

        LogActivity "`n$delimiter`nTarget Backup(s)`n$delimiter`n"
        LogActivity "Initiating Target DB Backup(s) ..."
        Invoke_SQLDBBackup $targetserver $target_backupdirectory $target_dbs 
        LogActivity "`nTarget backup completed successfully.`n" "Green"
        #start-sleep -s 45
        #LogActivity "   Target Database Backups are taken successfully to the location $target_backupdirectory" "CYAN"

        #LogActivity "Taking backups is completed successfully..." "Green"
        #Start-Sleep -s 2
     }
     catch
     {
        Write-Host "Error Ocurred during backups " -f Red
        Exit 0
     }

    #*****************copy the logins******************
    LogActivity "Scripting users, roles and permissions ..." -console_disabled 0 
    Copy_Logins -targetserver $targetserver -backuplocation $target_backupdirectory -targetdbs $target_dbs
    LogActivity "Scripted users, roles and permissions ..."
    LogActivity "`n" -console_disabled 0
    #Start-Sleep -s 2

    #*****************take the db properties********************
    LogActivity "Scripting the DB Properties ..." -console_disabled 0
    $target_db_properties = $target_db_properties[2]
    foreach($key in $target_db_properties.Keys)
    {
        $db_properties_file = "$target_backupdirectory\$key\target_db_properties.txt"
        New-Item -path $db_properties_file -ItemType File | Out-Null
        foreach($subkey in $target_db_properties[$key].Keys)
        {
            "$subkey : $($target_db_properties[$key][$subkey])" >> $db_properties_file
        }
        
        "`n****************************************************`n" >> $db_properties_file
        LogActivity "File Path: $db_properties_file" -console_disabled 0
    }
    
    LogActivity "Scripted the DB Properties."
    #Start-Sleep -s 2

    #------------------------------------------------------------***DB_REFRESH_IMPLEMENTATION***--------------------------------------------------
    LogActivity "`n$delimiter`nRestore`n$delimiter"
    $db_match_dictionary = [ordered]@{}
    $count = $source_dbs.count
    $index = 0
    while($index -lt $count)
    {
        $db_match_dictionary[$target_dbs[$index]] = $source_dbs[$index]
        $index++
    }

    $smoServer = New-Object Microsoft.SqlServer.Management.Smo.Server($targetserver)
    try 
    {
        foreach($db in $db_match_dictionary.Keys)
        {
            $global:targetDBName = $db
            $sourcedb = $db_match_dictionary[$db]
            $filepath = get_data_log_filepath $db
            $log_filepaths = @($filepath | where-object {($_.split("."))[-1] -contains "ldf"})
            if($log_filepaths.count -eq 0)
            {
                Write-Host "There is no log file with .ldf extension in the target db $db. Exiting" -f Red
                Exit 0
            }
            $diskpath = "$source_backupdirectory\$sourcedb\$sourcedb.bak"
            $query = "USE [master]
                                Alter database [$targetDBName] set Restricted_user with rollback immediate 
                                RESTORE DATABASE [$targetDBName] FROM  DISK = N'$diskpath' WITH FILE = 1,"
            
            
            if($remove_dbs_sorted.Value -contains $sourcedb)
            {
                if($multiple_file_details[0].Keys -contains $sourcedb)
                {                            
                    $data_logical_m = $multiple_file_details[0][$sourcedb] | Get-Unique             
                    $data_filepath = @($filepath | Where-Object{$_.contains(".mdf") -or $_.contains(".ndf")})               
                    foreach($data in $data_logical_m)
                    {         
                        $data_file = $data_filepath[$($data_logical_m.indexof($data))]
                        
                        if(!($data_file -eq $null) -and $($data_logical_m.indexof($data)) -eq 0)
                        {
                            $data_file = join-path (split-path $data_file -Parent) "$($db)_automation_$(get-date -f 'MMddyyyy').mdf"
                        }
                        elseif(!($data_file -eq $null))
                        {
                            $data_file = join-path (split-path $data_file -Parent) "$($db)_data$($data_logical_m.indexof($data))_automation_$(get-date -f 'MMddyyyy').ndf"
                        }
                        else
                        {
                            $data_file = $data_filepath[0]
                            $data_file = join-path (split-path $data_file -Parent) "$($db)_data$($data_logical_m.indexof($data))_automation_$(get-date -f 'MMddyyyy').ndf"
                        }
                        $query += "`nMOVE N'$data' TO N'$data_file',"
                    }
                }
                if($multiple_file_details[1].Keys -contains $sourcedb)
                {
                    $log_logical_m = $multiple_file_details[1][$sourcedb] | Get-Unique 
                    $log_filepath = @($filepath | Where-Object{$_.contains(".ldf")})
            
                    foreach($log in $log_logical_m)
                    {
                        $log_file = $log_filepath[$($log_logical_m.indexof($log))]

                        if(!($log_file -eq $null) -and $($log_logical_m.indexof($log)) -eq 0)
                        {
                            $log_file = join-path (split-path $log_file -Parent) "$($db)_automation_$(get-date -f 'MMddyyyy').ldf"
                        }
                        elseif(!($log_file -eq $null))
                        {
                            $log_file = join-path (split-path $log_file -Parent) "$($db)_log$($log_logical_m.indexof($log))_automation_$(get-date -f 'MMddyyyy').ldf"
                            
                        }
                        else
                        {
                            $log_file = $log_filepath[0]
                            $log_file = join-path (split-path $log_file -Parent) "$($db)_log$($log_logical_m.indexof($log))_automation_$(get-date -f 'MMddyyyy').ldf"
                        }
                        $query += "`nMOVE N'$log' TO N'$log_file',"
                    }
                }
                $query +=  "`nNOUNLOAD,  REPLACE,  STATS = 5
                            GO
                            Alter database [$targetDBName] set multi_user"
                
                $restore_query = $query
            }
            else 
            {
                $logicalname = get_logicalname $sourcedb "$source_backupdirectory\$sourcedb"
                $global:backupFile = "$source_backupdirectory\$sourcedb\$sourcedb.bak" 
                $restore_query = ConstructQuery $targetDBName $sourcedb $filepath $log_filepaths $logicalname $backupFile
            }
        
            LogActivity "`nQuery to execute the restore for the db $db `n$restore_query`n" -console_disabled 0
            LogActivity "`nInitiating restore for the db `'$db`'...."
            try
            {
                LogActivity "Start Time : $(get-date -Format "MM/dd/yyyy hh:mm:ss") EST"
                $restore_time = Measure-command {Invoke-Sqlcmd -ServerInstance $targetserver -Database $db -query $restore_query -ErrorAction stop}
                LogActivity "End Time :  $(get-date -Format "MM/dd/yyyy hh:mm:ss") EST"
                LogActivity "Restore of the db `'$db`' completed successfully" "Green"
                #LogActivity "`n    Time taken for the restore of the $db : $([Math]::Round($restore_time.TotalMinutes, 5)) minutes"
            }
            catch
            {
                Invoke-sqlcmd -serverinstance $targetserver -Database $db -Query "Alter database [$db] set multi_user"
                LogActivity "`nDB Refresh Failed. Please review and proceed with Manual Implementation" "Red"
                Exit 0
            }

            $change_owner = "Use `"$db`" `ngo`nsp_changedbowner 'sa'"
            try
            {
                Invoke-sqlcmd -serverinstance $targetserver -query $change_owner -ErrorAction Stop
            }
            catch
            {
                Write-Host "Error occurred in changing the owner"
            }
            
            $success_perms_log_path = "$local_workfolder\LoginPermissions_Log.txt"
            $failure_perms_log_path = "$local_workfolder\LoginPermissionsWarning_Log.txt"
            if(!(Test-Path $success_perms_log_path))
            {
                New-Item -path $success_perms_log_path -ItemType File | Out-Null
            }
            $success_set = 1
            $users = ($smoServer.Databases[$db].Users | Where-Object {!($_.Id -le 4)}).name
            foreach($user in $users)
            {
                try
                {
                    $smoServer.Databases[$db].Users[$user].drop()
                    "Successfully dropped the user $user for the db '$db'" >> $success_perms_log_path
                }
                catch
                {
                    if(!(Test-Path $failure_perms_log_path))
                    {
                        New-Item -path $failure_perms_log_path -ItemType File | Out-Null
                    }
                    "Failed to drop the user $user for the db '$db'. Kindly manual check" >> $failure_perms_log_path
                    $success_set = 0
                }
            }
            "`n$delimiter`n" >> $success_perms_log_path
            if(!($success_set))
            {
                "`n$delimiter`n" >> $failure_perms_log_path
                Write-Warning "Kindly check the user_drop logs at $($failure_perms_log_path)"
            }
        
            $loginsfile = "$target_backupdirectory\$db\$($db)_login.sql"
            try
            {
                $collation_property = "Collation"
                #Write-Host "Target Collation : $($target_db_properties[$db].$collation_property)"
                #Write-Host "Source Collation : $($source_db_properties[$sourcedb])"
                if($($target_db_properties[$db].$collation_property) -eq $source_db_properties[$sourcedb])
                {

                    LogActivity "Applying users, roles, and permissions..."
                    Invoke-Sqlcmd -ServerInstance $targetserver -Database $db -InputFile $loginsfile -ErrorAction Stop
                    LogActivity "Permissions script of the db `'$db`' applied successfully." "Green"
                }
                else
                {
                    LogActivity "The collation level of source and target dbs are different. Kindly proceed manually for applying permissions" 
                }

            }
            catch
            {
                LogActivity "Please Review Logs for any exceptions $failure_perms_log_path" -console_disabled 1
                "Error Occurred in restoring the login permissions. $_" >> $failure_perms_log_path
            }
            #LogActivity "`nRestoration of the logins, roles and permissions for the db $db happened successfully" "Green"
            #$image_index = @($db_match_dictionary.Keys).indexof($db)
            #$FilePath = "$local_workfolder\Image2_Implementation$image_index.png"
            #$bounds = [Drawing.Rectangle]::FromLTRB(0, 0, 2000, 2000)
            #TakeSnaps $bounds $FilePath
        }

        LogActivity "`n$delimiter`nImplementation of DB refresh is successful.`nTarget Server: $targetserver`nDatabase Name(s) : $($($target_dbs | foreach-object {"`'$_`'"}) -join ", ")`nCurrent Date & Time : $(get-date -f 'MM/dd/yyyy HH:mm:ss') EST`n$delimiter`n" "Green"
        #Start-Sleep -s 2
        #$FilePath = "$local_workfolder\Image3_Implementation_Completion.png"
        #$bounds = [Drawing.Rectangle]::FromLTRB(0, 0, 2000, 2000)
        #TakeSnaps $bounds $FilePath
    }
    catch 
    {
        Write-Host "Error Occurred in implementation : $_"
        Exit 0
    }

    #------------------------------------------------post_validation------------------------------------------------------
    if($target_env -ilike "dev*" -or $target_env -ilike "test*")
    {
        $recovery_query = ""
        foreach($db in $target_dbs)
        {
            $recovery_query += "Alter Database [$db] set recovery SIMPLE`n"
        }
        Invoke-Sqlcmd -serverinstance $targetserver -query $recovery_query -ErrorAction Stop
    }
    
    

    #$target_db_properties = $target_db_properties[2]
    #LogActivity "`n******Post-Implementation*****`n"
    #$properties_array = @("AutoCreateStatisticsEnabled", "Autoshrink", "AutoUpdateStatisticsEnabled", "Collation", "CompatibilityLevel", "IsFullTextEnabled", "RecoveryModel", "UserAccess", "Version", "Owner", "IsDatabaseSnapshot", "IsReadCommittedSnapshotOn", "SnapshotIsolationState")
    #$smo_target_server = New-Object Microsoft.SQLServer.Management.SMO.Server $targetserver
    #foreach($db in $target_dbs)
    #{
    #    LogActivity "Applying DB Properties for the $db..."
    #    $database = $smo_target_server.Databases[$db]
    #    $no_change = 0
    #    foreach($property in $properties_array)
    #    {
    #        if($($database.$property) -eq $target_db_properties[$db].$property)
    #        {
    #        
    #        }
    #        else 
    #        {
    #            $no_change = 1
    #            if($property -eq "Owner")
    #            {
    #                $newOwner = $target_db_properties[$db].$property
    #                $changeOwnerQuery = "Use `"$db`"
    #                                     go
    #                                     sp_changedbowner '$newOwner'"
    #                 Invoke-sqlcmd -ServerInstance $targetserver -query $changeOwnerQuery -ErrorAction Stop
    #                
    #            }
    #            else
    #            {
    #                $database.$property = $target_db_properties[$db].$property
    #            }
    #            
    #        }
    #    }
    #    try
    #    {
    #        if($no_change -eq 1)
    #        {
    #            $database.Alter()
    #        }
    #    }
    #    catch
    #    {
    #        
    #    }
    #}
    #LogActivity "`nDB Properties script applied successfully." "Green"

    
    #ConvertImagesTo-PDF -folder $local_workfolder -CR $CRNumber
    #$subjectname = $username.trim("Admin_")
    #TriggerMail $htmlBody $CRNumber $subjectname
    LogActivity "**************Post Implementation**************`n" 
    
    LogActivity "**Database Status Check**" 
    $post_check_target_smo = New-Object Microsoft.SQLServer.Management.SMO.Server $targetserver
    foreach($db in $target_dbs)
    {
        $dbDetails = $post_check_target_smo.Databases[$db]
        if($dbDetails.Status -eq [Microsoft.SqlServer.Management.Smo.DatabaseStatus]::Normal)
        {
            LogActivity "DB '$db' is online on the target server $targetserver" "Green" 
        }
        else 
        {
            LogActivity "DB '$db' is not online. Kindly make it online manually" "Red" 
        }
    }
    LogActivity "`n**Drive Space Check on Target Server**" -log_disabled 1
    foreach($key in $space_details.Keys)
    {
        LogActivity "Free space available in the drive $key : $([Math]::Round($($space_details[$key][0]),2)) GB ($([Math]::Abs($($($space_details[$key][1]))*100))%)" 
    }
    
    LogActivity "`nDetailed Log Path : $(hostname)`n$txt_logsfile" "Green" 
    LogActivity "`nCurrent Date & Time : $(get-date -f 'MM/dd/yyyy HH:mm:ss') EST"
    ConvertTo-PDF $logsfile
    LogActivity "`n$delimiter`nImplementation of DB refresh is successful.`nRITM : $RITM`nCR : $CRNumber`nTarget Server : $targetserver`nDatabase Name(s) : $($($target_dbs | foreach-object {"`'$_`'"}) -join ", ")`nCurrent Date & Time : $(get-date -f 'MM/dd/yyyy HH:mm:ss') EST`n$delimiter`n`nAttach the PDF document to respective Change Tasks" -log_disabled 1 -console_disabled 1 -htmlenabled 1


    $subjectname = $username -ireplace "Admin_",""
    #TriggerMail $htmlBody $CRNumber $subjectname
    Exit 1
    
}
catch 
{
    Write-Host "Error Occurred : $_"
}