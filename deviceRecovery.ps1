 <# 
Script Name: Device Recovery (Most Wanted)
Does: 1) Pull list of devices from FW that should be returned in the next 21 days.
      2) Twenty-one days before term date, emails the principal and tech to let them know.
      3) Schedule task.
#>

Clear-Host             

$today = $(get-date -f yyyy-MM-dd)
$homebase = split-path -parent $MyInvocation.MyCommand.Definition
$log = $homebase + '\mostwanted\' +  $(get-date -f yyyy-MM-dd) +'log.txt'
write-host $date
$time = get-date



#Create connection to SQL
$uid = "dbuser"
$pwd = "pwd"
$database = "dbname"
$server = "dbserver"

$connstring = "server=$server;uid=$uid;pwd=$pwd;database=$database;integrated security=false"

$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connstring

### WON'T YOU TAKE ME TO...FUNCTION TOWN

function Get-MostWanted {

$sqlcmd = New-Object System.Data.SqlClient.SqlCommand

$sqlcmd.CommandText = $query 

$sqlcmd.Connection = $connection

$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
$sqladapter.SelectCommand = $sqlcmd
$list = New-Object System.Data.DataSet
$rowcount = $sqladapter.Fill($list) 
$connection.Close()

return $list

}

function Set-MostWanted {

$sqlcmd = New-Object System.Data.SqlClient.SqlCommand

$sqlcmd.CommandText = $quandry 

$sqlcmd.Connection = $connection

$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
$sqladapter.SelectCommand = $sqlcmd
$list = New-Object System.Data.DataSet
$rowcount = $sqladapter.Fill($list) 
$connection.Close()

return $list

}

function Get-Buildings {

$sqlcmd = New-Object System.Data.SqlClient.SqlCommand

$sqlcmd.CommandText = "SELECT distinct entity, email FROM dbo.dbTable"

$sqlcmd.Connection = $connection

$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
$sqladapter.SelectCommand = $sqlcmd
$list = New-Object System.Data.DataSet
$rowcount = $sqladapter.Fill($list) 
$connection.Close()

return $list

}

#list of techs by building
function Get-Techs {

$sqlcmd = New-Object System.Data.SqlClient.SqlCommand

$sqlcmd.CommandText = $techquar 

$sqlcmd.Connection = $connection

$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
$sqladapter.SelectCommand = $sqlcmd
$list = New-Object System.Data.DataSet
$rowcount = $sqladapter.Fill($list) 
$connection.Close()

return $list

}


function Import-SQLPS {
    $Current = Get-Location
    Import-Module sqlps -DisableNameChecking
    Set-Location $Current
}


function Run-Usp {

$sqlcmd = New-Object System.Data.SqlClient.SqlCommand

$sqlcmd.CommandText = "EXEC usp_pullEmpInfo"

$sqlcmd.Connection = $connection

$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
$sqladapter.SelectCommand = $sqlcmd
$list = New-Object System.Data.DataSet
$rowcount = $sqladapter.Fill($list) 
$connection.Close()

return $list

}


#report on devices that will soon need to be recovered
Run-Usp

$buildapool = Get-Buildings

$buildings = $buildapool.Tables.Rows | Select-Object entity, email

 function Build-HTML
{

$madmenpool = Get-MostWanted $query

$madmen = $madmenpool.Tables.Rows | Select-Object device, authUser, termDate, deviceType, email, mac

   $b = '<html>'
  $b += '<head><title>DEVICE RECOVERY</title></head>'
  $b += '<body class="fixed-header"><iam-portal>'

  $b += '<style> .descript { font-style: italic  }
    table tr:nth-child(odd) { background-color:#dddddd } @supports (-webkit-overflow-scrolling: touch) { .update { display:none } }  @media (min-width: 600px) {
 .portlink {  } .update { background-color:green; color: white; width: 15%; padding: 4px; display: block}  li { display: inline; margin-left: 2px } a, a:visited { text-decoration: none;background-color: #0C3F0A; color: white; padding: 2px } a:hover { background-color: #2E54FF }    </style>'

 
  $b += '<table width="100%"; border="0"; cellspacing="0"; cellpadding="0">'
    
  #  $b += '<tr style="width:33%"; id="stumismatch"><h2>DEVICES</h2></tr>'
    $b += '<thead>'
    $b += '<tr style="background-color:blue; color:white;text-align: left"><th>device</th><th>employee</th><th>term date</th><th>device type</th></tr>'
    $b += '</thead>'
   

### DEVICE BY DEVICE

$b += '<tbody>'
if ($madmen) {
 $b += "<p>Dear Administrator and Building Technician,</p><p>The employee(s) listed below has a termination date in the past or within the next 14 calendar days. Our records indicate that this employee currently has an iPad or laptop issued to them as a work device. Please make arrangements to collect the device on or before their last day of work.</p><p> If you have questions about this email, please contact the Help Desk. Thank you.</p>"
foreach ($r in $madmen) 
{
  write-host $r 
$device = $r.device
$mac = $r.mac 
  $userton = $r.authUser
  $date = $r.termDate 
  $type = $r.deviceType
  
  $evalor = $device + $userton + $date + $type 
  write-host $evalor
    
   
$b += "<tr><td>" + $device + "</td><td>" + $userton + "</td><td>" +  $date + "</td><td>" + $type + "</td></tr>"

}

  $b += '</table>'
   
$b += '</tbody>'


    $b += '</body></html>'
     
    return $b


} # end of if madmen
    ###



}


foreach ($b in $buildings) {
$entity = $b.entity
$email = $b.email
write-host "email is $email"


$query = "SELECT DISTINCT DEVICE, TERMDATE, DEVICETYPE, ENTITY, EMAILSENT, EMAIL, AUTHUSER FROM dbTable where emailSent = '$today' and entity = '$entity'"
$techquar = $null
$techquar = "SELECT techMail FROM dbo.buitechmap where ent = '$entity'"

$techapool = Get-Techs
$techos = $null
$techos = $techapool.Tables.Rows | Select-Object ent, techMail
$thetech = $null
$thetech = $techos.techMail

write-host "the entity for this time and place is $entity"

    $from = "Information Services <skycheck@mccsc.edu>"
    $us = "arrayofISemailaddresses"
    $smtp = 'fqdnsmtpserver'
    $subject = "Device Recovery"
    
    $body = Build-HTML
    write-host $body
    write-host "an email will be sent to $email now"
if ($body) {
Send-MailMessage -To $email -From "MCCSC Information Services $from" -cc $thetech -bcc $us -Subject $subject -BodyAsHtml -Body $body -SmtpServer $smtp -dno OnFailure
$outing = "email sent to $email with tech $thetech"
$outing | out-file $log -append 
}
else { write-host "nothing could be found that hasn't already been found."}
                      # }
}
 
