#=============================================
# Author: Fabian Raygosa
# Date: 8/24/2012
# Purpose: Collect size of site collections
#   Email attachment information
#   Put data into database for recording purposes
#=============================================

#Set Thread to clean memory
$ver = $host | select version
if ($ver.Version.Major -gt 1)  {$Host.Runspace.ThreadOptions = "ReuseThread"}

#Check for snapin
if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null) 
{
	Add-PSSnapin Microsoft.SharePoint.PowerShell;
}

# Open SQL connection (you have to change these variables)
# ------------------- DB CONN SETUP ------------------------------------------
$DBServer = "{DB SERVER}}"
$DBName = "{DB NAME}"
$DBUser = "{USER}"
$DBPass = '{PASSWORD}}'
#Currently only MSSQL
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$sqlConnection.ConnectionString = "Server=$DBServer;Database=$DBName;User ID=$DBUser;Password=$DBPass;"
$executeSQL = 0
# ------------------- DB CONN SETUP ------------------------------------------
$LogFile = "D:\Tasks\PowershellAutoScripts\sitecolsize.html"
#SMTP server name
$smtpServer = "{EMAIL SERVER}"
$emailFrom = "{EMAIL FROM}"
#comma seperated
$emailTo = "{Email To}"
$emailSubject = "Powershell Auto: Sharepoint Site Collection Sizes"
$emailMessage = "See Attachment"

write-output "" > $LogFile

Get-SPSite -Limit All | select url, @{label="Size in MB";Expression={$_.usage.storage/1MB}} | Sort-Object -Descending -Property "Size in MB" | ConvertTo-Html -title "Site Collections sort by size" | Set-Content $LogFile

# Quit if the SQL connection didn't open properly.
# -------------------- OPEN DB -----------------------
if($executeSQL -eq 1)
{
    $sqlConnection.Open()
    if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
        "Connection to DB is not open."
        Exit
    }
}

#($_.disksizerequired/1024/1024)
#Get-SPSite -Limit All | select url, @{label="Size in MB";Expression={$_.usage.storage/1MB}} | Sort-Object -Descending -Property "Size in MB"
$DBCol = Get-SPSite -Limit All | Sort-Object -Descending -Property "Size"
foreach($line in $DBCol) {
    $Url = $line.url
    $Size = ($line.usage.storage/1MB)
    if( $Size -ge 1000 ) {
        #Write-Host "Name: " $line.Name
        #Write-Host "Size: " $Size
        if($executeSQL -eq 1) {
            $sqlCommand = New-Object System.Data.SqlClient.SqlCommand
            $sqlCommand.Connection = $sqlConnection
            # This SQL query will insert 1 row based on the parameters, and then will return the ID
            # field of the row that was inserted.
            $gbsize = $Size / 1000
            $sqlString = "insert into spStatsSiteColSize(SCUrl, SCSizeMB, SCSizeGB) values('$Url', $Size, ($Size/1000));"
            $sqlCommand.CommandText = $sqlString
            #Write-Host $sqlString
            # Run the query and get the scope ID back into $InsertedID
            $InsertedID = $sqlCommand.ExecuteScalar()
        }
    }
}

# -------------- Close the connection ----------------------.
if ($sqlConnection.State -eq [Data.ConnectionState]::Open) {
    $sqlConnection.Close()
} 
# -------------- Close the connection ----------------------

Start-Sleep -s 2

#Creating a Mail object
$msg = new-object Net.Mail.MailMessage

#Creating SMTP server object
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

#Email structure 
$msg.From = $emailFrom
$msg.ReplyTo = $emailFrom
$msg.To.Add($emailTo)
$msg.subject = $emailSubject
$msg.body = $emailMessage

$att = New-Object Net.Mail.Attachment($LogFile, 'text/html')
$msg.Attachments.Add($att)

#Sending email 
$smtp.Send($msg)
$att.Dispose();