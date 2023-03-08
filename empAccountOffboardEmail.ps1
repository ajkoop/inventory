 <# Author: Andrew Koop
   Purpose: Let outgoing employees know that they don't get to keep their accounts forever :)
   What It Do:
    1) Run stored procedure to grab employees with term date in the next 22 days
    2) Email those individuals
    3) Log the send
#>
Clear-Host
$today = $(get-date -f yyyy-MM-dd)
$homebase = split-path -parent $MyInvocation.MyCommand.Definition
$log = $homebase + '\accountOffboardEmails\' +  $(get-date -f yyyy-MM-dd) +'log.txt'

#Create connection to SQL
$uid = "dbu"
$pwd = "pwd"
$database = "dbname"
$server = "server"

$connstring = "server=$server;uid=$uid;pwd=$pwd;database=$database;integrated security=false"

$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connstring

function Build-HTML
{
    
    $b = '<html>'
    $b += '<body>'
   $b += "<p>Greetings, $fname $lname.</p><p>Our records indicate that you have an upcoming exit date of $termdate. Please note that your access to MCCSC accounts (mccsc.net, mccsc.edu, etc.) will expire at the end of your last day of employment. Thank you.</p><p><Sincerely, </p><p>MCCSC Information Services</p>"
    $b += '</body>'
    $b += '</html>'

    return $b

}


#pulls data from account provisioning database and financial SIS
function Run-Usp {

$sqlcmd = New-Object System.Data.SqlClient.SqlCommand

$sqlcmd.CommandText = "EXEC usp_empDeputize"

$sqlcmd.Connection = $connection

$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
$sqladapter.SelectCommand = $sqlcmd
$list = New-Object System.Data.DataSet
$rowcount = $sqladapter.Fill($list) 
$connection.Close()

return $list

}


Run-Usp

#send to principals and building technicians
function Send-Message ($s)
{
    
    $from = 'intaddress@mccsc.edu'
    #$from = $testmail
    $smtp = 'fqdnsmtpserv'
    $subject = 'MCCSC Account Offboarding on ' + $termdate 
    $body = Build-HTML

    #$to1,$to2,$to3,$to4,$to5|
    Send-MailMessage -To "$fname $lname <$email>" -From $from -Subject $subject -BodyAsHtml -Body $body -SmtpServer $smtp
}

function Get-StaffData($query)
{
    $sqlcmd = New-Object System.Data.SqlClient.SqlCommand
    $sqlcmd.CommandText = $query
    $sqlcmd.Connection = $connection

    $sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $sqladapter.SelectCommand = $sqlcmd
    $dataset = New-Object System.Data.DataSet
    $sqladapter.Fill($dataset)
    $connection.Close()

    return $dataset
}

#Set daily log

$query = "SET NOCOUNT ON; SELECT username, termDate, fname, lname FROM dbo.dbTable where emailSent = '$today'"
$dataset = Get-StaffData($query)
$rowcount = $dataset.Tables[0].Rows.Count

$data = $dataset.Tables[0]


foreach($d in $data)
{
   $staff = "" | SELECT username, termDate, fname, lname
   $email = $d.username + '@mccsc.edu'
   
   $fname = $d.fname
   $lname = $d.lname
   $term = $d.termdate 
   $termdate = $term.ToString("M-d-yyyy")
   
   
    Send-Message $s

    $strOut = "$fname $lastname was emailed about term date on $termdate"
    
    $strOut | Out-File $log -Append
}
 
