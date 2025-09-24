$input_file = "D:\Pavithra\SQL_Release_Check\input.txt"
$input_content = get-content $input_file
$sql_versions_hashtable = @{}
foreach($line in $input_content)
{
    $sql_versions_hashtable[$line.split(":")[0].trim()] += $line.split(":")[1].trim()
}

$release_url = "https://learn.microsoft.com/en-us/troubleshoot/sql/releases/download-and-install-latest-updates"
$response = Invoke-WebRequest -Uri $release_url 
$tables = $response.ParsedHtml.getElementsByTagName('Table')

# get the list of versions available
$version_Table = $tables[0]
$rows = $version_Table.getElementsByTagName('tr')  # Get all rows
$version = @()
foreach ($row in $rows) 
{
    $cells = $row.getElementsByTagName("td")  # Get all columns in the row
    if($cells -ne $null)
    {
        $version += $($cells[0].innerText.split("`n")[0].split(" ")[-1]).trim()
        $index += 1
    }
}

$versions = $($version | where-object {$_.trim() -notin @("2012","R2","2008")}) | Where-Object {$_.trim()}
Write-Host "`nAvailable SQL Versions : " -nonewline 
Write-Host "$($version -join ',')" -f Cyan

#get the latest available service/cumulative update from each version
$available_sql_versions_hashtable = [ordered]@{}
foreach($version in $versions)
{
    $sql_version_table = $tables[$versions.indexof($version)+1]
    $rows = $sql_version_table.getElementsByTagName('tr')  # Get all rows
    
    foreach ($row in $rows) 
    {
        $cells = $row.getElementsByTagName("td")  # Get all columns in the row
        $values = @()       
        foreach ($cell in $cells) 
        {
            $values += $cell.innerText.Trim()  # Extract and clean cell text
        }
        if(($values[2] -notmatch "Azure") -and $values[2] -ne $null)
        {
            Write-Host "The latest available CU for the SQL $version is $($values[0])"#$values[2]
            $available_sql_versions_hashtable[$version] += @($($values[0]), $($values[2]), $($values[3]), $($values[4]))
            break
        }        
    }
}

#comparision of the current and available sql releases
$release_set = 0
$new_release_kbnumber = @()
$keys = @($available_sql_versions_hashtable.Keys)
foreach($key in $keys)
{
    if(@($sql_versions_hashtable.Keys) -contains $key.trim())
    {
        $available_build_number = $available_sql_versions_hashtable[$key][0].split(".")[-2].trim()
        $current_build_number = $sql_versions_hashtable[$key.trim()].split(".")[-2].trim()
        if($available_build_number -eq $current_build_number)
        {
            Write-Host "No release for the version $key"
            $available_sql_versions_hashtable[$key] += @("0")
        }
        else 
        {
            $release_set = 1
            Write-Host "There is a new release for version $($key.trim())"
            Write-Host "Build Number : $($available_sql_versions_hashtable[$key][0])"
            Write-Host "Release Date : $($available_sql_versions_hashtable[$key][3])"
            $available_sql_versions_hashtable[$key] += 1
            $new_release_kbnumber += $available_sql_versions_hashtable[$key][2]     
        }
    }
    else 
    {
        $release_set = 1
        Write-Host "There is a new release for version $($key.trim())"
        Write-Host "Build Number : $($available_sql_versions_hashtable[$key][0])"
        Write-Host "Release Date : $($available_sql_versions_hashtable[$key][3])"
        $available_sql_versions_hashtable[$key] += 1
        $new_release_kbnumber += $available_sql_versions_hashtable[$key][2] 
    }
}

#set content to the input file
Clear-Content -path $input_file
foreach($key in $available_sql_versions_hashtable.Keys)
{
    "$key : $($available_sql_versions_hashtable[$key][0])" >> $input_file
}

#download the kbs
$output_path = "C:\Users\PKathirv\OneDrive - JNJ\Desktop\Documents\EngActivities\SQL ReleaseCheck"
$HTML_Content = "<br>"
foreach($key in $keys)
{
    if($available_sql_versions_hashtable[$key][-1] -eq 1)
    {
        $kbnumber = $available_sql_versions_hashtable[$key][-3]
        $link_to_route = $($response.Links | where-object {$_.innertext -eq "$kbnumber"})[0]
        if($link_to_route.'data-linktype' -like 'relative')
        {
            $link_to_route = $($response.Links | where-object {$_.innertext -eq "$kbnumber"})[0].href
            $link_to_route = $(Split-Path $release_url) + "/" + $link_to_route
            $link_to_route = $link_to_route.replace("\","/")
            $required_response = Invoke-WebRequest -uri $link_to_route
            $download_link = $($required_response.Links | Where-Object {$_.innerhtml -match "Download the latest cumulative update package"})[0].outerhtml.split('"')[1]
            $download_response = Invoke-WebRequest -uri $download_link
            $download_patch = $($download_response.Links | Where-Object {$_.innerhtml -match "<span>Download</span>"})[0].href
        }
        #elseif($link_to_route.'data-linktype' -like 'external')
        #{
        #    $link_to_route = $($response.Links | where-object {$_.innertext -eq "$kbnumber"})[0].href
        #    $link_to_route = $link_to_route.replace("\","/")
        #    $required_response = Invoke-WebRequest -uri $link_to_route
        #    $download_link = $($required_response.Links | Where-Object {$_.innerhtml -match "Download the package now"})[0].outerhtml.split('"')[1]
        #    $download_response = Invoke-WebRequest -uri $download_link
        #    $download_patch = $($download_response.Links | Where-Object {$_.innerhtml -match "<span>Download</span>"})[0].href
        #}

        $destination = Join-Path $output_path "SQL$key"
        if(!(Test-Path -path $destination))
        {
            New-Item $destination -ItemType Directory
        }
        #Write-Host "Downloading the patches to the DML Lcations....."
        ##Invoke-WebRequest -Uri $download_patch -OutFile "$destination\$($available_sql_versions_hashtable[$key][0]).exe"
        ##Write-Host "Downloaded the patch $($available_sql_versions_hashtable[$key][0]) for $key to the location $destination\$($available_sql_versions_hashtable[$key][0]).exe"
        #$HTML_Content += "Downloaded the patch $($available_sql_versions_hashtable[$key][0]) for $key to the location $destination\$($available_sql_versions_hashtable[$key][0]).exe<br>"
    }

}


# trigger mail
$HTML = '<style type="text/css">
#Header{font-family:"Calibri", Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;}
#Header td, #Header th {font-size:12px;border:1px solid #98bf21;padding:3px 7px 2px 7px;}
#Header th {font-size:14px;text-align:left;padding-top:5px;padding-bottom:4px;background-color:#C0C0C0;color:#fff;}
#Header tr.alt td {color:#000;background-color:#EAF2D3;}
</Style>'

$HTML += "<HTML><BODY><Table border=1 cellpadding=0 cellspacing=0 id=Header>
<TR>
    <TH><B>SQL Release</B></TH>
    <TH><B>Latest Build Number</B></TH>
    <TH><B>Update</B></TH>
    <TH><B>KB Number</B></TH>
    <TH><B>Release Date</B></TH>
    <TH><B>New Release?</B></TH>
</TR>"

foreach($key in $available_sql_versions_hashtable.Keys)
{
    if($available_sql_versions_hashtable[$key][-1] -eq 1)
    {
        $HTML += "<TR style='background-color: white; color: green;'>"
        $answer = "YES"
        $bold_tag = "<B>"
        $close_bold_tag = "</B>"
    }
    else
    {
        $HTML += "<TR style='background-color: white; color: black;'>"
        $answer = "NO"
        $bold_tag = ""
    }
    
    $HTML += "<TD>$($bold_tag)$($key)$($close_bold_tag)</TD>
              <TD>$($bold_tag)$($available_sql_versions_hashtable[$key][0])$($close_bold_tag)</TD>
              <TD>$($bold_tag)$($available_sql_versions_hashtable[$key][1])$($close_bold_tag)</TD>
              <TD>$($bold_tag)$($available_sql_versions_hashtable[$key][2])$($close_bold_tag)</TD>
              <TD>$($bold_tag)$($available_sql_versions_hashtable[$key][3])$($close_bold_tag)</TD>
              <TD>$($bold_tag)$($answer)$($close_bold_tag)</TD>
            </TR>"
}
$HTML += "</TABLE>"
if($release_set)
{
    $to = "PKathirv@its.jnj.com", "sborkar3@its.jnj.com", "skuma554@its.jnj.com", "HShah39@ITS.JNJ.com"
    #$cc = "nobody@its.jnj.com"
    $cc = "juppala@its.jnj.com", "APande59@ITS.JNJ.com", "svijaya7@ITS.JNJ.com", "mmangala@its.jnj.com"
    #$HTML += "<br>@PKathirv@its.jnj.com please check there is a new release"
    $HTML += $HTML_Content
    $HTML += "</Body></HTML>"
    $subject = "MSSQL Release Status : New release is available"
}
else 
{
    $to = "PKathirv@its.jnj.com"
    $cc = "sborkar3@its.jnj.com", "skuma554@its.jnj.com"
    $subject = "MSSQL Release Status : NO New release"
    $HTML += $HTML_Content
    $HTML += "</Body></HTML>"
}
$emailParams = @{
    From       = "SA-NCSUS-SQLHELPDESK@its.jnj.com"
    To         = $to
    CC         = $cc
    Subject    = $subject
    Body       = $HTML
    BodyAsHtml = $true
    SmtpServer = "smtp.na.jnj.com"
}

Send-MailMessage @emailParams
