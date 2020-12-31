# ================================================================================================================
# |Developer: Fabian Raygosa
# |Date: September 20, 2012
# |Purpose: This script will scan through a sharepoint farm sites and then get data on who has alerts created on the farm.
# html file is built plus an email can be sent
# Scan for UNCOMMENT's
# ================================================================================================================

#Set Thread to clean memory
$ver = $host | select version
if ($ver.Version.Major -gt 1)  {$Host.Runspace.ThreadOptions = "ReuseThread"}

#Check for snapin
if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null) 
{
	Add-PSSnapin Microsoft.SharePoint.PowerShell;
}

#Get All the Sites
$site = Get-SPSite -Limit All

#Set vars
$loopc = 0
$truecount = 0
$message = ""
$path_to_file = 'D:\{Path}\{TO}}\{HTML FILE}.htm'
$smtpServer = "{EMAIL SERVER HERE}"
$from_email = "{EMAIL}"
$to_email = "{TO EMAIL COMMA SEPERATED}"
$email_subject = "Powershell Auto: Sharepoint Alerts in site farm"
$email_message = "See attachment"

#Empty Output for html File UNCOMMENT
write-output "" > $path_to_file

#Loop through sites and built html
foreach($webName in $site)
{
	$loopc = 0
	$web = Get-SPWeb $webName.URL
	foreach($alert in $web.Alerts)
	{
		$count++
		if($loopc -eq 0)
		{
			$loopc = 1
			$message += "<tr bgcolor=red><td colspan=6 align=center><b>SITE: <u>" + $webName.URL + "</u></b></td></tr>"
		}
		if($alert.User.Name.Length -gt 5)
		{
			$truecount++
		}
		#	"User         - " + $alert.User.Name
		#	"Title        - " + $alert.Title
		#	"Frequency    - " + $alert.AlertFrequency
		#	"Delivery Via - " + $alert.DeliveryChannels
		#	"Change Type  - " + $alert.eventtype
		#	Write-Host "_____________________"

		$message += "<tr><td><font color=black><b>&nbsp;" + $alert.User.Name + "</b></font></td>"
		$message += "<td><font color=green><b>" + $alert.Title + "</b></font></td>"
		$message += "<td><font color=blue><b>" + $alert.AlertFrequency + "</b></font></td>"
		$message += "<td><font color=orange><b>" + $alert.DeliveryChannels + "</b></font></td>"
		$message += "<td><font color=grey><b>" + $alert.eventtype + "</b></font></td>"
		$message += "<td><font color=grey><b>" + $alert.ListUrl + "</b></font></td></tr>"
	}
}

$t = date
$front = "<table border=1><tr><td colspan=6 align=center><b>"+ $t +"</b><br><b>TOTAL ALERTS: "+ $count +" TOTAL NON SYSTEM ALERTS: " + $truecount + "</b></td></tr><tr BGCOLOR='#99CCFF'><td align=center><b>User</b></td><td align=center><b>Title</b></td><td align=center><b>Frequency</b></td>"
$front += "<td align=center><b>Delivery Via</b></td><td align=center><b>Change Type</b></td><td align=center><b>List Url</b></td></tr>"

$message = $front + $message

$message += "</table>"
write-output $message >> $path_to_file

Start-Sleep -s 2

#Creating a Mail object
$msg = new-object Net.Mail.MailMessage

#Creating SMTP server object
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

#Email structure 
$msg.From = $from_email
$msg.ReplyTo = $from_email
$msg.To.Add($to_email)
$msg.subject = $email_subject
$msg.body = $email_message

$att = New-Object Net.Mail.Attachment($path_to_file, 'text/html')
$msg.Attachments.Add($att)

#Sending email 
$smtp.Send($msg)
$att.Dispose();
