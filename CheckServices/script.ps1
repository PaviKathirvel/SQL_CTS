
function TriggerMail
{

    param($htmlfile, $time, $timezone)

    $htmlBody = Get-Content -Path $htmlfile -Raw

    $emailParams = @{
        From       = "SA-NCSUS-SQLHELPDESK@its.jnj.com"
        To         = "PKathirv@its.jnj.com"
        Subject    = "DB Services Check - $time $timezone"
        Body       = $htmlBody
        BodyAsHtml = $true
        SmtpServer = "smtp.na.jnj.com"
    }

    Send-MailMessage @emailParams

}


$time = get-date -f "MMM dd, yyyy hh:mm"
$timezone = ((get-timezone).id -split " " | % {$_[0]}) -join ""
$formattedtime = (get-date $time).tostring("MMddyyyhhmm")
$OutputFile = "D:\Service_Check\SQLServiceStatus_$($formattedtime).htm"
$ServerList = Get-Content "D:\Service_Check\serverlist.txt"

$Result = @()
Foreach($ServerName in $ServerList)
{
    $ServicesStatus = get-wmiobject -ComputerName $ServerName -Class win32_service | 
                      where {$_.name -like '*SQL*' -and $_.startmode -like '*Auto*'} | where {$_.name -inotlike "*SQLTELEMETRY*" -and $_.name -inotlike "*SQLWRITER*"} -ErrorAction SilentlyContinue
	$start_time = $(Invoke-Sqlcmd -serverinstance $ServerName -Database "master" -Query "SELECT sqlserver_start_time FROM sys.dm_os_sys_info;").sqlserver_start_time
	$Result += New-Object PSObject -Property @{
	    ServerName = $ServerName
		ServiceName = $ServicesStatus.name
		Status = $ServicesStatus.State
        StartTime = $start_time
	}
}

if($Result -ne $null)
{
	$HTML = '<style type="text/css">
	#Header{font-family:"Calibri", Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;}
	#Header td, #Header th {font-size:12px;border:1px solid #98bf21;padding:3px 7px 2px 7px;}
	#Header th {font-size:14px;text-align:left;padding-top:5px;padding-bottom:4px;background-color:#C0C0C0;color:#fff;}
	#Header tr.alt td {color:#000;background-color:#EAF2D3;}
	</Style>'

    $HTML += "<HTML><BODY><Table border=1 cellpadding=0 cellspacing=0 id=Header>
		<TR>
			<TH><B>Server Name</B></TH>
            <TH><B>SQL Server Boot Time</B></TH>
			<TH><B>Service Name</B></TD>
			<TH><B>Status</B></TH>
		</TR>"

    $HTML_failed = ""
    $HTML_passed = ""
    Foreach($Entry in $Result)
    {
        if($Entry.Status -ne "Running")
		{
			$HTML_failed += "<TR bgColor=Red>"
            $HTML_failed += "
						<TD>$($Entry.ServerName)</TD>
                        <TD>$($Entry.StartTime)</TD>
						<TD>$($($Entry.ServiceName) -join ', ')</TD>
						<TD>$($Entry.Status)</TD>
					</TR>"
		}
		else
		{
			$HTML_passed += "<TR>"
		    $HTML_passed += "
						<TD>$($Entry.ServerName)</TD>
                        <TD>$($Entry.StartTime)</TD>
						<TD>$($($Entry.ServiceName) -join ', ')</TD>
						<TD>$($Entry.Status)</TD>
					</TR>"
		}

    }
    $HTML += $HTML_failed + $HTML_passed + "</Table></BODY></HTML>"

	$HTML | Out-File $OutputFile
}

TriggerMail -htmlfile $OutputFile -time $time -timezone $timezone