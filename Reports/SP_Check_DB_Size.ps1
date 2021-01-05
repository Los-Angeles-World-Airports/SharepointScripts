#=============================================
# Purpose: List All databases in a farm with their size.
#       Save to a file and email to someone
#       Save to a database (helpful for reporting)
#=============================================

#Set Thread to clean memory
$ver = $host | select version
if ($ver.Version.Major -gt 1)  {$Host.Runspace.ThreadOptions = "ReuseThread"}

#Check for snapin
if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null) 
{
	Add-PSSnapin Microsoft.SharePoint.PowerShell;
}


# ----------------------------SET STETINGS--------------------------
$thresholdDBsize = 1000
# ------------------- DB CONN SETUP ------------------------------------------
$DBServer = "{SQL SERVER}"
$DBName = "{DATABASE NAME}"
$DBUser = "{User}"
$DBPass = '{Password}'
#set your connection example - currently set for MS SQL
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$sqlConnection.ConnectionString = "Server=$DBServer;Database=$DBName;User ID=$DBUser;Password=$DBPass;"
#If you do not want to do run sql set to true
$executeSQL = 0
# ------------------- DB CONN SETUP ------------------------------------------
#SMTP server name
$smtpServer = "{Email Server}"
$emailFrom = "{EMAIL FROM}"
$emailTo = "{Email TO - comma seperated}"
$emailSubject = ""
$emailMessage = ""
$sendEmail = 0
#logFile
$LogFile = "{PATH\TO\LOGFILE.txt}"
# ----------------------------SET STETINGS--------------------------

#get database
$message = Get-SPDatabase | Sort-Object disksizerequired -desc | Format-Table Name, @{Label ="Size in MB"; Expression = {$_.disksizerequired/1024/1024}} 

# Quit if the SQL connection didn't open properly.
# -------------------- OPEN DB -----------------------
if($executeSQL -eq 1) {
    $sqlConnection.Open()
    if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
        "Connection to DB is not open."
        Exit
    }
}
# -------------------- OPEN DB -----------------------
#($_.disksizerequired/1024/1024)
#Get-SPDatabase | Sort-Object disksizerequired -desc | foreach-object{insert_into_db($_.Name,[System.Int64] $_.disksizerequired) } 
$DBSet = Get-SPDatabase | Sort-Object disksizerequired -desc
foreach($line in $DBSet) {
    $Name = $line.Name
    $Size = ($line.disksizerequired/1024/1024)
    if( $Size -ge $thresholdDBsize ) {
        #Write-Host "Name: " $line.Name
        #Write-Host "Size: " $Size
        $sqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $sqlCommand.Connection = $sqlConnection
        # This SQL query will insert 1 row based on the parameters, and then will return the ID
        # field of the row that was inserted.
        $gbsize = $Size / 1000
        $sqlString = "insert into spStatsDBSize(DBName, DBSizeMB, DBSizeGB) values('$Name', $Size, ($Size/1000));"
        $sqlCommand.CommandText = $sqlString
        #Write-Host $sqlString

        # Run the query and get the scope ID back into $InsertedID
        if($executeSQL -eq 1) {
            $InsertedID = $sqlCommand.ExecuteScalar()
        }
    }
}
# -------------- Close the connection ----------------------.
if($executeSQL -eq 1) {
    if ($sqlConnection.State -eq [Data.ConnectionState]::Open) {
        $sqlConnection.Close()
    } 
}
# -------------- Close the connection ----------------------

Write-Output $message | Out-File $logFile

if($sendEmail -eq 1) {

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

    $att = New-Object Net.Mail.Attachment($LogFile, 'text/plain')
    $msg.Attachments.Add($att)

    #Sending email 
    $smtp.Send($msg)
    $att.Dispose();
}
