#######################################################################################
#
#  Script Name : Get_Win32_Product                                             
#  Overview    : Getting computer installation program list (Using [Get-WmiObject])   
#  Argument    : None                                                    
#  Create Date : 2018-05-18
#  Auther      : nag
#
#######################################################################################

####<Pre>###################################################################################

#Get Output Filename Info

## Get Desktop Path
$userprofile=$env:userprofile

## Get Timestamp
$tstamp = Get-Date -Format "yyyy-MMdd-HHmmss"

## Set Output Filename
$out_filename=$userprofile+ "\Desktop\" + $tstamp + "_application_list.csv"

# Set Server List File

## Get Server List File-Name
$server_list="server.txt"

## Input information to be processed for Server List-file
$f=Import-Csv $server -Delimiter "," -Header "svname","cred","pdat"

## Row Count Server List File
$reccnt=$(Get-Content server99.txt| Measure-Object).Count


####<Main>##############################################################################

#Output Header Record
Write-Output "`"Server Name`",`"Aplication NAme`",`"Vendor`",`"Version`""  | Out-File $out_filename -Encoding default -Append

# Loop Until the Server List-file Ends
$i=1
$linecnt=0
 

foreach ($l in $f) {
$prog_percent = $linecnt / $reccnt *100
Write-Progress "Get Information" ([String]$prog_percent + "%") -percentComplete $prog_percent

$linecnt=$linecnt + 1

## Set Computername for Server List Record
$svname=$l.svname

## Judgment Credential Type & Set to Credential

### Case of Domain Administrator
if ($l.cred ="da") {
   $credential="mydomain\Administrator"
}
else {

### Case of Local Administrator
 if ($L.cred="la") {
  $credential=$l.svname + "\Administrator"
 }
}


### Set Crypt Password File-Name
$pdfilename =$l.pdat + ".dat"

### Get Password String(SecureString) 
$password = Get-Content $pdfilename | ConvertTo-SecureString

### Create New Credential Object
$credential_obj = New-Object System.Management.Automation.PsCredential $credential, $password

### Set Hostname String
$HOSTNAME = $svname

### Exec [Get-WmiObject]

  try {


       $res_str=Get-WmiObject -Class Win32_Product -ComputerName $HOSTNAME -Credential $credential_obj | Select Name,Vendor,Version |ConvertTo-Csv -NoTypeInformation | select -skip 1 

      }catch [Exception]{

       echo "Error " + $HOSTNAME

     }
### Loop Until Object Records Ends
foreach ($line in $res_str) 
{

#### Join Hostname $ Object Record  
 $outrec= "`"" + $HOSTNAME + "`"," + $line
 
#### Output to the CSV file
 Write-Output $outrec  | Out-File $out_filename -Encoding default  -Append 
 
 }

$res_str=""

$i++

}
