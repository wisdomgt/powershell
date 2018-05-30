 #################################################################################
 #
 #  Script Name : Get_Install_List_From Registory                                 
 #  Overview    : Getting Computer Installed Program List (Using [Get-WmiObject]) 
 #  Argument    : None                                                    
 #  Create Date : 2018-05-24
 #  Auther      : nag
 #
 #################################################################################
 
 
 Function Get-AppReg ([int64]$HKEY,[string]$REG,[int]$remoteflg,[string]$HOSTNAME,$credential) {
 
 #*******************************************************************************
 # <Function Get-AppReg
 #  Argument
 #   1 HKEY(Int64)  HKEY_CLASSES_ROOT   :2147483648 
 #                  HKEY_CURRENT_USER   :2147483649 
 #                  HKEY_LOCAL_MACHINE  :2147483650 
 #                  HKEY_USERS          :2147483651 
 #                  HKEY_CURRENT_CONFIG :2147483653 
 #                  HKEY_DYN_DATA       :2147483654
 #
 #   2 REG(String)  "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
 #                  "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
 #
 #   3 remoteflg(Int)  Localhost:0 RemoteHost:1
 #
 #   4 HOSTNAME(String) Specify Name resolvable Computername
 #   5 credential(PsCredential) Credential Object
 #*******************************************************************************
 
 # Set (NameSpace,ListType,Hostname)
   $ns = "root\default"
   $cls = "StdRegProv"
 
 # Create Result Object(Array) 
   $result = New-Object System.Collections.ArrayList
 
 # Getting WmiObject
 
 
   if ($remoteflg -eq 1) { 
     try {
             $wmi = Get-WmiObject -List $cls -Namespace $ns -ComputerName $HOSTNAME -Credential $credential 
         }catch [Exception]{
             echo "Error Connect NameSpace" 
         }
 
   }else{
     try {
             $wmi = Get-WmiObject -List $cls -Namespace $ns -ComputerName $HOSTNAME 
          }catch [Exception]{
             echo "Error Connect NameSpace"
         }
   }
 
 # Get Registory SubKey 
 
   try {
 
     $r=$wmi.EnumKey($HKEY, $REG) 
 
   }catch [Exception]{
     echo "Error Get EnumKey" 
   }
 
 # Loop Until Sub Key Is Empty 
 for ($i=0 ;$i -lt $r.sNames.Count; $i++) {
 
      $regkey2=  $REG + "\" + $r.sNames[$i]
 
 　　　try {
      $r2=$wmi.EnumValues($HKEY, $regkey2) 
 
      }catch [Exception]{
       echo "Error GetEnumValues" 
      }
 
      $obj = New-Object -TypeName PSObject
 
      Add-Member -InputObject $obj -MemberType NoteProperty -Name ComputerName -Value $HOSTNAME
      Add-Member -InputObject $obj -MemberType NoteProperty -Name RegKey -Value $r.sNames[$i]     
 
       for($j = 0; $j -lt $r2.sNames.Count; $j++) {
 
          $sub_progress=$j/$r2.sNames.Count*100
          $sub_progress_str="サブキー走査状況( " + $regkey2 + " )" 
          Write-Progress $sub_progress_str ([String]$sub_progress.ToString("0.00") + "%") -percentComplete $sub_progress -Id 2
 
         if    (($r2.sNames[$j] -eq "DisplayName")    `
            -or ($r2.sNames[$j] -eq "DisplayVersion") `
            -or ($r2.sNames[$j] -eq "InstallDate")    `
            -or ($r2.sNames[$j] -eq "InstallDate")    `
            -or ($r2.sNames[$j] -eq "Publisher"))     `
            {
 
              try {
                switch ($r2.Types[$j]) {
                    1 { # REG_SZ
                        $val = $wmi.GetStringValue($HKEY, $regkey2, $r2.sNames[$j])
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name $r2.sNames[$j] -Value $val.sValue 
                    }
                    2 { # REG_EXPAND_SZ
                        $val = $wmi.GetExpandedStringValue($HKEY, $regkey2, $r.sNames[$j])
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name $r2.sNames[$j] -Value $val.sValue 
                    }
                    3 { # REG_BINARY
                        $val = $wmi.GetBinaryValue($HKEY, $regkey2, $r2.sNames[$j])
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name $r2.sNames[$j] -Value $valdname.uValue
                    }
                    4 { # REG_DWORD
                        $val = $wmi.GetDWORDValue($HKEY, $regkey2, $r2.sNames[$j])
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name $r2.sNames[$j] -Value $val.uValue
                    }
                    7 { # REG_MULTI_SZ
                        $val = $wmi.GetMultiStringValue($HKEY, $regkey2, $r2.sNames[$j])
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name $r2.sNames[$j] -Value $val.sValue
                    }
                    11 { # REG_QWORD
                        $val = $wmi.GetQWORDValue($HKEY, $regkey2, $r2.sNames[$j])
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name $r2.sNames[$j] -Value $val.uValue
                    }
                    default { # Invalid Object
                        #Add-Member -InputObject $obj -MemberType NoteProperty -Name $r2.sNames[$j] -Value "Null"
                    }
                }
              }catch [Exception]{
                     echo "Error i="$i "j="$j + $HOSTNAME + $regkey2 + $r2.sNames[$j]
 
              }
         }
 
     }
 
       if ($obj.DisplayName) {
 
                  if ($obj.DisplayName.Contains("Update for Microsoft")) {
                      # No Action
                  }else{
                      $result.Add($obj) | Out-Null
                  }
        }
 }     
 
    return $result 
 
 #return $result |ConvertTo-Csv
 
 }
 
 ################################################################################
 # Main Proc
 #  Vriable Set and Call Function(Get-AppReg)
 ################################################################################
 
 
 # Variable
 ##------------------------------------------------------------------------------
 ## Set Server List File-Name
 ## List File Format(Hostname:String],CredentialType:String,CryptFileName:String)
 ## Please Change Your ListFiles
 ## ex. $server_list="xxxxxx.txt"  
 ##------------------------------------------------------------------------------
 
 $server_list="server99.txt"
 
 $fpath=".\"+$server_list
 if (Test-Path $fpath) {
 
 }else{
   $server_list=""
 }
 
 ##------------------------------------------------------------------------------
 ## If ServerLIST does not exist, set variable for local mode
 ##------------------------------------------------------------------------------
 
 if ($server_list -eq ""){   
     $remote_flg=0
      $localhost=[Net.Dns]::GetHostName()
      $lfstr=$localhost+",xx,xx"
      $f=ConvertFrom-Csv $lfstr -Delimiter "," -Header "svname","cred","pdat"
      $reccnt=1
 
 }else{
 
  ## Input information to be processed for Server List-file
     $f=Import-Csv $server_list -Delimiter "," -Header "svname","cred","pdat"
     $remote_flg=1
  ## Row Count Server List File
     $reccnt=$(Get-Content $server_list| Measure-Object).Count
 
 }
 
 #Set Output Filename Info
 ## Get Desktop Path
 $userprofile=$env:userprofile
 
 ## Get Timestamp
 $tstamp = Get-Date -Format "yyyy-MMdd-HHmmss"
 
 ## Set Output Filename
 $out_filename=$userprofile+ "\Desktop\" + $tstamp + "_application_list.csv"
 
 # Loop Until the Server List-file Ends
 $i=1
 $linecnt=0
 
 ##------------------------------------------------------------------------------
 ## Loop until ServerLIST is empty
 ##------------------------------------------------------------------------------
 
    foreach ($l in $f) {
 
        ## Set Computername for Server List Record
        $svname=$l.svname
 
        $prog_percent = $linecnt / $reccnt *100
        $curr_state="進捗状況"+"( " +$svname+ " )" 
        Write-Progress $curr_state ([String]$prog_percent.ToString("0.00") + "%") -percentComplete $prog_percent -Id 1
 
        $linecnt=$linecnt + 1
 
 
 
        ## Judgment Credential Type & Set to Credential
 
        if ($remote_flg -eq 1) {
 
            if ($l.cred -eq "da") {               ### Case of Domain Administrator
                $credential="mydomain\Administrator"
            }else{
                if ($l.cred -eq "la") {           ### Case of Local Administrator
                    $credential=$l.svname + "\Administrator"
                }
            }
 
 
        ### Set Crypt Password File-Name
                $pdfilename =$l.pdat + ".dat"
 
        ### Get Password String(SecureString) 
                $password = Get-Content $pdfilename | ConvertTo-SecureString
 
        ### Create New Credential Object
                $credential_obj = New-Object System.Management.Automation.PsCredential $credential, $password
        }
        ### Set Hostname String
               $HOSTNAME = $svname
 
        ### Exec [Get-WmiObject]
 
 
 
 ##------------------------------------------------------------------------------
 ## Exec Function and Output Csv File
 ##------------------------------------------------------------------------------
 
 
        Get-AppReg 2147483650 "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" $remote_flg $HOSTNAME $credential_obj|select computername,DisplayName,DisplayVersion,Publisher,InstallDate | Export-Csv $out_filename -Append -Encoding Default 
        Get-AppReg 2147483650 "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" $remote_flg $HOSTNAME $credential_obj|select computername,DisplayName,DisplayVersion,Publisher,InstallDate | Export-Csv $out_filename -Append -Encoding Default 
        Get-AppReg 2147483649 "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" $remote_flg $HOSTNAME $credential_obj|select computername,DisplayName,DisplayVersion,Publisher,InstallDate  | Export-Csv $out_filename -Append -Encoding Default 
 
 
  }
