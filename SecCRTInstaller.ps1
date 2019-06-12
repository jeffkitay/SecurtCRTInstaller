function Clean-CerticateStore(){
    $ErrorActionPreference = 'silentlycontinue'
    $CertsSmartCard,$CertsSmartCardNot,$CertstokeepSAN,$CertstokeepSANNot,$CertstokeepSANOverflow,$CertsExpired,$CertsExpiredNot,$CertsAffialiteA,$CertsAffialiteANot = Get-UserCertificates
    try{
        $RemoveExtraCertsmartCard2 = $CertsExpired
        $catch = $RemoveExtraCertsmartCard2 | Remove-Item -Force ##  #-WhatIf
        #Expired certs
        $Stop='' 
    }
    Catch{
      Throw
    }      
}
Function Copy-Ini{

    param
    (
        [Parameter(Mandatory = $false, Position=0)]
        [String]$Param=''
    )
    try{
        $RS = Test-Path -Path "$scriptDirectory\files\NCBI PIV Remote Access.ini"
                If ($RS){
        $Filecontent = Get-Content "$scriptDirectory\files\NCBI PIV Remote Access.ini"
            }Else    {
        $Filecontent = Is-ConfigFile
    }
        $Catch = New-Item -ItemType Directory -Force -Path $configSessionPath -Confirm:$false
        $Catch = Copy-Item "$ScriptDirectory\Files\Config" -Destination $configSession -Recurse -force -Confirm:$false
        $Outfile = "$configSessionPath" + '\' + "MSLOGIN01 Unconfigured NCBI PIV Remote Access.ini"
        $Trap = Out-File -FilePath $Outfile -Force -InputObject $Filecontent -Confirm:$false
        ##
        $GoodCertOnly=@()
        $Certs = Compare-object -ReferenceObject @($CertsSmartCard | Select-Object)  -DifferenceObject @($CertsAffialiteA | Select-Object) -PassThru -Property Thumbprint -IncludeEqual ##
                if (-not $CertsExpired){
        $CertsExpired = New-Object -TypeName psobject
    }
                        foreach ($C in $Certs){
        if (($CertsExpiredNot -contains $C) -and ($c.subject -match "OU=NIH" ) -and ($c.subject -match '-a')){
            $Goodcertonly = $GoodCertOnly + $C
        }    
    }
        ## If Multiple Certs
                                                        if($Goodcertonly.Count -gt 1){
        foreach ($Choice in $CertsAffialiteA){
            $button = 'OK' # OK only; https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.interaction.msgbox?view=netframework-4.7.2
            $title = "There Are Multiple Valid PIV Certificates`n"
            $message = "`nSelect the appropriate certificate`n`n$($Choice.Subject)`nValidity Period `n$($Choice.NotBefore) `nthrough `n$($Choice.NotAfter)"    
            $Returnvalue=Show-Messagebox -message $message -title $title -timeout '120' -buttonset 'yn' -icontype 'exclamation'
            If (($Returnvalue -eq 1) -or ($Returnvalue -eq 6) -or ($Returnvalue -eq -1) ){
                $ChoiceCert= $Choice
                break
            }
        }
        
    }
                If (-not $ChoiceCert){
      $ChoiceCert = $GoodCertOnly[0]
    }
        $ncbipcname=Show-Inputbox -message "Enter NCBI PC desktop." -title "NCBI Desktop Name" -default ""
                if ($ncbipcname){
        $PCCONFIG = $Filecontent -replace 'mslogin01', $ncbipcname  # | Set-Content "$configSessionPath\$sshSessionName"        
    }
                foreach ($G in $ChoiceCert){
        $SN = ($($G.serialnumber) -replace '(..)','$1 ').trim(' ')
        $MSLOGIN = $Filecontent -replace "AB CD EF GH", $SN}
            $Outfile = "$configSessionPath" + '\' + "MSLOGIN01 NCBI PIV Remote Access.ini"
            $Trap = Out-File -FilePath $Outfile -Force -InputObject $MSLOGIN -Confirm:$false
                            If ($ncbipcname){
            $PCCONFIG= $PCCONFIG -replace "AB CD EF GH", $SN
            $outfile = "$configSessionPath" + '\' + "$($ncbipcname.ToUpper()) PIV Remote $sshSessionName"    
            Out-File -FilePath $outfile -Force -InputObject $PCCONFIG -Confirm:$false
        }    
            Write-Debug 'End Copy Ini'  
            return     
    }
    Catch{
        Throw
    }   
}
function Copy-License{
    [CmdletBinding()]
    [OutputType([boolean])]
    param 
    (
    [string]$program = ''
    )
    Try{
        $LicenseExist = Test-Path -Path "$scriptDirectory\files\SecureCRT_SecureFX.lic"
        $License = "$scriptDirectory\files\SecureCRT_SecureFX.lic"
        $EXE = Test-Path -Path "C:\Program Files\VanDyke Software\Clients\securecrt.exe"
        $DEST = "C:\Program Files\VanDyke Software\Clients"
                if ($LicenseExist -and $EXE){
      $RS = Copy-Item -Path $License -Destination $dest -Force 
    }
                                        Else{
      $Files = gci -Path "$scriptDirectory\files" -include '*.lic' -Recurse -file | Select-Object -First 1
      $Dest = gci -Path "$env:ProgramFiles" -Include 'securecrt.exe' -Recurse | Select-Object -First 1 
        If ($false -eq $Dest){
            $Dest = gci -Path ${env:ProgramFiles(x86)} -Include 'securecrt.exe' -Recurse | Select-Object -First 1            
        } 
      $dest = $dest.DirectoryName
      $RS = Copy-Item -Path $($files.FullName) -Destination $dest -Force
    }
   
        Return $true
    }
    Catch{
        Throw
    } 
}
   Function Get-MsiProperties{
     [OutputType([hashtable])]
     param
     (
        [Parameter(Mandatory = $true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path $_})]
        #[ValidateScript({$_.EndsWith(".msi")})]
        [String]$MsiPath=$files
     )

     try{
   
       $MsiPath = Resolve-Path $MsiPath
       #Create the type
       $HashProperties = @{}
       $type = [Type]::GetTypeFromProgID("WindowsInstaller.Installer")
       $installer = [Activator]::CreateInstance($type)
       #write-host "installer = $installer"
       #write-host "msipath = $MsiPath"

       #The OpenDatabase method of the Installer object opens an existing database or creates a new one, returning a Database object
       #For our case, we need to open the database in read only. The open mode is 0
       $db = Get-MsiPropertiesInvokeMemberOnType "OpenDatabase" $installer @($MsiPath,0)
   
       #The OpenView method of the Database object returns a View object that represents the query specified by a SQL string.
       $view = Get-MsiPropertiesInvokeMemberOnType "OpenView" $db ('SELECT * FROM Property')

       #The Execute method of the View object uses the question mark token to represent parameters in an SQL statement.
       Get-MsiPropertiesInvokeMemberOnType "Execute" $view $null

       #The Fetch method of the View object retrieves the next row of column data if more rows are available in the result set, otherwise it is Null.
       $record = Get-MsiPropertiesInvokeMemberOnType "Fetch" $view $null
       while($record -ne $null)
       {
         $property = Get-MsiPropertiesInvokeMemberOnType "StringData" $record 1 "GetProperty"
         $value = Get-MsiPropertiesInvokeMemberOnType "StringData" $record 2 "GetProperty"
         Write-debug "Property = $property : Value = $value"
         if ($property){
            $Hashproperties.Add($property,$value)
         }
         $record = Get-MsiPropertiesInvokeMemberOnType "Fetch" $view $null
       }

       #The Close method of the View object terminates query execution and releases database resources.
       Get-MsiPropertiesInvokeMemberOnType "Close" $view $null
       return $HashProperties
     }
     Catch{
      Throw
     }
}
Function Get-MsiPropertiesInvokeMemberOnType{
    param
    (
        #The string containing the name of the method to invoke
        [Parameter(Mandatory=$true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$Name,
       
        #A bitmask comprised of one or more BindingFlags that specify how the search is conducted.
        #The access can be one of the BindingFlags such as Public, NonPublic, Private, InvokeMethod, GetField, and so on
        [Parameter(Mandatory=$false, Position = 3)]
        [System.Reflection.BindingFlags]$InvokeAttr = "InvokeMethod",

        #The object on which to invoke the specified member
        [Parameter(Mandatory=$true, Position = 1)]
        [ValidateNotNull()]
        [Object]$Target,

        #An array containing the arguments to pass to the member to invoke.
        [Parameter(Mandatory=$false, Position = 2)]
        [Object[]]$Arguments = $null
    )

    Try{
      $Target.GetType().InvokeMember($Name,$InvokeAttr,$null,$Target,$Arguments)
    }
    Catch{
      Throw
    }
}
Function Get-ObjectProperty {
  <#
      .SYNOPSIS
      Get a property from any object.
      .DESCRIPTION
      Get a property from any object.
      .PARAMETER InputObject
      Specifies an object which has properties that can be retrieved.
      .PARAMETER PropertyName
      Specifies the name of a property to retrieve.
      .PARAMETER ArgumentList
      Argument to pass to the property being retrieved.
      .EXAMPLE
      Get-ObjectProperty -InputObject $Record -PropertyName 'StringData' -ArgumentList @(1)
      .NOTES
      This is an internal script function and should typically not be called directly.
      .LINK
      http://psappdeploytoolkit.com
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$true,Position=0)]
    [ValidateNotNull()]
    [object]$InputObject,
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateNotNullorEmpty()]
    [string]$PropertyName,
    [Parameter(Mandatory=$false,Position=2)]
    [object[]]$ArgumentList
  )
	
  Begin { }
  Process {
    ## Retrieve property
    Try{
      Write-Output -InputObject $InputObject.GetType().InvokeMember($PropertyName, [Reflection.BindingFlags]::GetProperty, $null, $InputObject, $ArgumentList, $null, $null, $null)
    }
    Catch{
      Throw
    }
  }
  End { }
}
function Get-UserCertificates{
    [CmdletBinding()]
    [OutputType([array])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false)] 
        $Param1,

        # Param2 help description
        [Parameter(Mandatory=$false)]        
        [int]
        $Param2,

        # Param3 help description
        [Parameter(Mandatory=$false)]
        [String]
        $Param3
    )
    ##
    $Certstore = Get-ChildItem 'cert:\CurrentUser\My' | sort -unique
    $CertsExpiredNot= $Certstore | ? {(Get-Date $_.NotAfter) -gt (Get-Date)}
    $CertsAffialiteA = $Certstore | ? {$_.subject -match '-a' -and $_.subject -match 'OID.0.9.2342.19200300.100.1' -and $_.subject -notmatch 'ms-org'}
    $CertsExpired = Compare-Object -ReferenceObject @($Certstore | Select-Object)  -DifferenceObject @($CertsExpiredNot | Select-Object)  -PassThru -Property Thumbprint ##    
    $CertsAffialiteANot= Compare-Object -ReferenceObject @( $Certstore | Select-Object) -DifferenceObject @($CertsAffialiteA | Select-Object) -PassThru -Property Thumbprint ##
    ##
    $Certlistfilter=""
    $CertstokeepSAN=@()
    $CertstokeepSANNot=@()
    $CertstokeepSANOverflow=@()
    foreach ($Cert in $Certstore){    
        if ($sanExt= $cert.Extensions | Where-Object {$_.Oid.FriendlyName -match "subject alternative name"}){
            $sanObjs = new-object -ComObject X509Enrollment.CX509ExtensionAlternativeNames
            $altNamesStr=[System.Convert]::ToBase64String($sanExt.RawData)
            $sanObjs.InitializeDecode(1, $altNamesStr)
            $Certlistfilter =""
            Foreach ($SAN in $sanObjs.AlternativeNames){
                if ($Certlistfilter){break}
                $SAN = $SAN.strValue
                if ($SAN){
                    $CertListFilter = $SAN | ? {$_ -match '@nih.gov' -and $_ -notmatch '\$@' -and $Cert.subject -match '-a' -and $Cert.subject -match '-OID.0.9.2342.19200300.100.1'  } # keep ncbi upn
                }
                if ($CertListFilter){
                    $CertstokeepSAN = $Certstokeep + $Cert
                    break
                }
                if ($SAN){
                    $CertListFilter = $SAN | ? {(($_ -match '.gov' -or $Cert.subject -match 'ou=nih') -and ($_-match '\$@' -or $_ -notmatch $env:USERNAME) -and (($Cert.subject -match '-a' -or $Cert.subject -match '-e'-or $Cert.subject -match '-s') -and ($Cert.subject -match 'OID.0.9.2342.19200300.100.1'))  -or $Cert.subject -match 'serialnumber=' )} #remove $ and -a or -e from store
                }
                if ($CertListFilter){
                    $CertstokeepSANNot = $CertstokeepSANNot + $Cert
                    break
                }
                if ($SAN -and !$CertListFilter){
                    $CertListFilter = $SAN
                }
                if ($CertListFilter){
                    $CertstokeepSANOverFlow = $CertstokeepSANOverFlow+ $Cert
                    break
                }
            }
        }
     }
    ##
    $PrevErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'silentlycontinue'
    $Certlistfilter=""
    $CertsSmartCard=@()
    $CertsSmartCardNot=@()
    # Go through every certificate in the current user's "My" store
    $matched = $false
    foreach($Cert in $Certstore){
        $matched = $false
        foreach($extension in $Cert.Extensions){
            if ($matched){
                Break
            }
        # For each extension, go through its Enhanced Key Usages
            foreach($certEku in $extension.EnhancedKeyUsages){
                if ($matched){
                    Break
                }
                if($certEku.friendlyname -match "Smart Card Logon"){
                    $CertsSmartCard= $CertsSmartCard + $Cert
                    $matched=$true
                    Break
                }
            }            
                        
        }
    }
    $ErrorActionPreference = $PrevErrorActionPreference
    $CertsSmartCardNot =  Compare-Object -ReferenceObject  @($Certstore| Select-Object)  -DifferenceObject @($CertsSmartCard | Select-Object)  -PassThru -Property Thumbprint ##
    $ReturnObject=@()
    $ReturnObject = @($CertsSmartCard,$CertsSmartCardNot,$CertstokeepSAN,$CertstokeepSANNot,$CertstokeepSANOverflow,$CertsExpired,$CertsExpiredNot,$CertsAffialiteA,$CertsAffialiteANot)
    [int]$Counter='0'
    Foreach ($obj in $ReturnObject){
    [Array]$ReturnObject[$counter]= $obj | sort -unique
    $Counter++
    }
    $Catch = Return $ReturnObject
}
function Install-Application{
    [CmdletBinding()]
    [OutputType([boolean])]
    param 
    (
    [string]$program = 'SecureCRT',
    [Version]$version = '8.5.0',
    [boolean]$install=$true
    )
    $Files = gci -Path "$scriptDirectory\files" -include '*.msi' -Recurse -file
        foreach ($f in $files){
                [string]$string = $f.FullName
                $Params = "/i ""$string""",'/passive','/norestart',"/l*v ""$string.log""",'ALLUSERS=0'
                $p = Start-Process 'msiexec.exe' -ArgumentList $params -Verb RunAs -PassThru
                $count= 0
                while ($false -eq $p.HasExited){
                    Start-Sleep -Seconds 5
                    If ($count -ge 100){
                        Break
                    }
                    Else{
                    $Count++
                    }
                }
        }
    Return $true
}
Function Is-ConfigFile{
  [CmdletBinding()]
    [OutputType([string])]    
  Param (
    [Parameter(Mandatory=$false,Position=0)]
    [ValidateNotNull()]
    [object]$InputObject,
    [Parameter(Mandatory=$false,Position=1)]
    [ValidateNotNullorEmpty()]
    [string]$PropertyName,
    [Parameter(Mandatory=$false,Position=2)]
    [object[]]$ArgumentList
  )
  $DefaultIni=@'
S:"Username"=
S:"Password V2"=
S:"Login Script V3"=02:69ed0d0044bfb68ab8e3b851eeb862e99806502e56eb5d9295733b1fbe04e693b2517707a3e96ac2d76eaee570bf3cf9
D:"Session Password Saved"=00000000
S:"Local Shell Command Pre-connect V2"=02:69ed0d0044bfb68ab8e3b851eeb862e99806502e56eb5d9295733b1fbe04e6939f793a97f579e75c8f42ad70901ca0d4
S:"Monitor Username"=
S:"Monitor Password V2"=02:69ed0d0044bfb68ab8e3b851eeb862e99806502e56eb5d9295733b1fbe04e6935160e909e7f8861f5e7ff6a75d974d16
B:"Normal Font v2"=00000060
 f3 ff ff ff 00 00 00 00 00 00 00 00 00 00 00 00 90 01 00 00 00 00 00 01 00 00 00 01 76 00 74 00
 31 00 30 00 30 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00
 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 61 00 00 00
B:"Narrow Font v2"=00000060
 f3 ff ff ff 00 00 00 00 00 00 00 00 00 00 00 00 90 01 00 00 00 00 00 01 00 00 00 01 76 00 74 00
 31 00 30 00 30 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00
 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 61 00 00 00
D:"Use Narrow Font"=00000000
S:"SCP Shell Password V2"=
S:"PGP Upload Command V2"=02:69ed0d0044bfb68ab8e3b851eeb862e99806502e56eb5d9295733b1fbe04e693447429348d8b697ad3d20dd651fd8861
S:"PGP Download Command V2"=02:69ed0d0044bfb68ab8e3b851eeb862e99806502e56eb5d9295733b1fbe04e6935080e0df56e71c548ba44b64215c0212
D:"Is Session"=00000001
S:"Protocol Name"=SSH2
D:"Request pty"=00000001
S:"Shell Command"=
D:"Use Shell Command"=00000000
D:"Force Close On Exit"=00000000
D:"Forward X11"=00000001
S:"XAuthority File"=
S:"XServer Host"=127.0.0.1
D:"XServer Port"=00001770
D:"XServer Screen Number"=00000000
D:"Enforce X11 Authentication"=00000001
D:"Request Shell"=00000001
D:"Max Packet Size"=00001000
D:"Pad Password Packets"=00000001
S:"Sftp Tab Local Directory V2"=
S:"Sftp Tab Remote Directory"=
S:"Hostname"=ssh.ncbi.nlm.nih.gov
S:"Firewall Name"=None
D:"Allow Connection Sharing"=00000000
D:"Disable Initial SFTP Extensions"=00000000
D:"[SSH2] Port"=00000016
S:"Keyboard Interactive Prompt"=assword
S:"Key Exchange Algorithms"=gss-group1-sha1-toWM5Slw5Ew8Mqkay+al2g==,gss-gex-sha1-toWM5Slw5Ew8Mqkay+al2g==,diffie-hellman-group-exchange-sha1,diffie-hellman-group14-sha1,diffie-hellman-group1-sha1
D:"Use Global Host Key Algorithms"=00000001
S:"Host Key Algorithms"=ssh-rsa,ssh-ed25519,ecdsa-sha2-nistp256,ecdsa-sha2-nistp384,ecdsa-sha2-nistp521,null,x509v3-sign-rsa,x509v3-ssh-rsa,x509v3-sign-dss,x509v3-ssh-dss,ssh-dss
S:"Cipher List"=aes256-ctr,aes192-ctr,aes128-ctr,aes256-cbc,aes192-cbc,aes128-cbc,twofish-cbc,blowfish-cbc,3des-cbc,arcfour
S:"MAC List"=hmac-sha1,hmac-sha1-96,hmac-md5,hmac-md5-96
S:"SSH2 Authentications V2"=publickey
S:"Compression List"=none
D:"Compression Level"=00000005
D:"GEX Minimum Size"=00000400
D:"GEX Preferred Size"=00000800
D:"Use Global Public Key"=00000000
S:"Identity Filename V2"=
D:"Public Key Type"=00000001
D:"Public Key Certificate Store"=00000000
S:"PKCS11 Provider Dll"=
S:"Public Key Certificate Serial Number"=AB CD EF GH
S:"Public Key Certificate Issuer"=C=US, O=U.S. Government, OU=HHS, OU=Certification Authorities, CN=HHS-FPKI-Intermediate-CA-E1
S:"Public Key Certificate Username"=
D:"Use Username From Certificate"=00000000
D:"Certificate Username Location"=00000000
D:"Use Certificate As Raw Key"=00000000
S:"GSSAPI Method"=auto-detect
S:"GSSAPI Delegation"=full
S:"GSSAPI SPN"=host@$(HOST)
D:"SSH2 Common Config Version"=00000006
D:"Enable Agent Forwarding"=00000002
D:"Transport Write Buffer Size"=00000000
D:"Transport Write Buffer Count"=00000000
D:"Transport Receive Buffer Size"=00000000
D:"Transport Receive Buffer Count"=00000000
D:"Sftp Receive Window"=00000000
D:"Sftp Maximum Packet"=00000000
D:"Sftp Parallel Read Count"=00000000
D:"Preferred SFTP Version"=00000000
S:"Port Forward Filter"=allow,127.0.0.0/255.0.0.0,0 deny,0.0.0.0/0.0.0.0,0
S:"Reverse Forward Filter"=allow,127.0.0.1,0 deny,0.0.0.0/0.0.0.0,0
D:"Port Forward Receive Window"=00000000
D:"Port Forward Max Packet"=00000000
D:"Port Forward Buffer Count"=00000000
D:"Port Forward Buffer Size"=00000000
D:"Packet Strings Always Use UTF8"=00000000
D:"Auth Prompts in Window"=00000000
S:"Transfer Protocol Name"=SFTP
S:"Initial Directory"=
D:"Synchronize File Browsing"=00000000
D:"ANSI Color"=00000001
D:"Color Scheme Overrides Ansi Color"=00000001
S:"Emulation"=Xterm
D:"Enable Xterm-256color"=00000000
S:"Default SCS"=B
D:"Use Global ANSI Colors"=00000001
B:"ANSI Color RGB"=00000040
 00 00 00 00 a0 00 00 00 00 a0 00 00 a0 a0 00 00 00 00 a0 00 a0 00 a0 00 00 a0 a0 00 c0 c0 c0 00
 80 80 80 00 ff 00 00 00 00 ff 00 00 ff ff 00 00 00 00 ff 00 ff 00 ff 00 00 ff ff 00 ff ff ff 00
D:"Keypad Mode"=00000000
D:"Line Wrap"=00000001
D:"Cursor Key Mode"=00000000
D:"Newline Mode"=00000000
D:"Enable 80-132 Column Switching"=00000001
D:"Ignore 80-132 Column Switching When Maximized or Full Screen"=00000000
D:"Enable Cursor Key Mode Switching"=00000001
D:"Enable Keypad Mode Switching"=00000001
D:"Enable Line Wrap Mode Switching"=00000001
D:"Enable Alternate Screen Switching"=00000001
D:"WaitForStrings Ignores Color"=00000000
D:"SGR Zero Resets ANSI Color"=00000001
D:"SCO Line Wrap"=00000000
D:"Display Tab"=00000000
S:"Display Tab String"=
B:"Window Placement"=0000002c
 2c 00 00 00 00 00 00 00 01 00 00 00 fc ff ff ff fc ff ff ff fc ff ff ff fc ff ff ff 00 00 00 00
 00 00 00 00 00 00 00 00 00 00 00 00
D:"Is Full Screen"=00000000
D:"Rows"=00000018
D:"Cols"=00000050
D:"Scrollback"=000001f4
D:"Resize Mode"=00000000
D:"Sync View Rows"=00000001
D:"Sync View Cols"=00000001
D:"Horizontal Scrollbar"=00000002
D:"Vertical Scrollbar"=00000002
S:"Color Scheme"=Monochrome
S:"Output Transformer Name"=Default
D:"Use Unicode Line Drawing"=00000001
D:"Blinking Cursor"=00000001
D:"Cursor Style"=00000000
D:"Use Cursor Color"=00000000
D:"Cursor Color"=00000000
D:"Foreground"=00000000
D:"Background"=00ffffff
D:"Bold"=00000000
D:"Map Delete"=00000000
D:"Map Backspace"=00000000
S:"Keymap Name"=Xterm
S:"Keymap Filename V2"=
D:"Use Alternate Keyboard"=00000000
D:"Emacs Mode"=00000000
D:"Emacs Mode 8 Bit"=00000000
D:"Preserve Alt-Gr"=00000000
D:"Jump Scroll"=00000001
D:"Minimize Drawing While Jump Scrolling"=00000000
D:"Audio Bell"=00000001
D:"Visual Bell"=00000000
D:"Scroll To Clear"=00000001
D:"Close On Disconnect"=00000000
D:"Clear On Disconnect"=00000000
D:"Scroll To Bottom On Output"=00000001
D:"Scroll To Bottom On Keypress"=00000001
D:"CUA Copy Paste"=00000000
D:"Use Terminal Type"=00000000
S:"Terminal Type"=
D:"Use Answerback"=00000000
S:"Answerback"=
D:"Use Position"=00000000
D:"X Position"=00000008
D:"X Position Relative Left"=00000000
D:"Y Position"=00000008
D:"Y Position Relative Top"=00000000
D:"Local Echo"=00000000
D:"Strip 8th Bit"=00000000
D:"Shift Forces Local Mouse Operations"=00000001
D:"Ignore Window Title Change Requests"=00000000
D:"Copy Translates ANSI Line Drawing Characters"=00000000
D:"Copy to clipboard as RTF and plain text"=00000000
D:"Translate Incoming CR To CRLF"=00000000
D:"Dumb Terminal Ignores CRLF"=00000000
D:"Use Symbolic Names For Non-Printable Characters"=00000000
D:"Show Chat Window"=00000002
D:"User Button Bar"=00000002
S:"User Button Bar Name"=Default
S:"User Font Map V2"=
S:"User Line Drawing Map V2"=
D:"Hard Reset on ESC c"=00000000
D:"Ignore Shift Out Sequence"=00000000
D:"Enable TN3270 Base Colors"=00000000
D:"Use Title Bar"=00000000
S:"Title Bar"=
D:"Show Wyse Label Line"=00000000
D:"Send Initial Carriage Return"=00000001
D:"Use Login Script"=00000000
D:"Use Script File"=00000000
S:"Script Filename V2"=
S:"Script Arguments"=
S:"Upload Directory V2"=
S:"Download Directory V2"=
D:"XModem Send Packet Size"=00000000
S:"ZModem Receive Command"=rz\r
D:"Disable ZModem"=00000000
D:"ZModem Uses 32 Bit CRC"=00000000
D:"Force 1024 for ZModem"=00000000
D:"ZModem Encodes DEL"=00000001
D:"ZModem Force All Caps Filenames to Lower Case on Upload"=00000001
D:"Send Zmodem Init When Upload Starts"=00000000
S:"Log Filename V2"=
S:"Custom Log Message Connect"=
S:"Custom Log Message Disconnect"=
S:"Custom Log Message Each Line"=
D:"Log Only Custom"=00000000
D:"Generate Unique Log File Name When File In Use"=00000001
D:"Log Prompt"=00000000
D:"Log Mode"=00000000
D:"Start Log Upon Connect"=00000000
D:"Raw Log"=00000000
D:"Log Multiple Sessions"=00000000
D:"New Log File At Midnight"=00000000
D:"Trace Level"=00000000
D:"Keyboard Char Send Delay"=00000000
D:"Use Word Delimiter Chars"=00000000
S:"Word Delimiter Chars"=
D:"Idle Check"=00000000
D:"Idle Timeout"=0000012c
S:"Idle String"=
D:"Idle NO-OP Check"=00000001
D:"Idle NO-OP Timeout"=0000000f
D:"AlwaysOnTop"=00000000
D:"Line Send Delay"=00000005
D:"Character Send Delay"=00000000
D:"Wait For Prompt"=00000000
S:"Wait For Prompt Text"=
D:"Wait For Prompt Timeout"=00000000
D:"Send Scroll Wheel Events To Remote"=00000000
D:"Position Cursor on Left Click"=00000000
D:"Highlight Reverse Video"=00000001
D:"Highlight Bold"=00000000
D:"Highlight Color"=00000000
S:"Keyword Set"=<None>
S:"Ident String"=
D:"Raw EOL Mode"=00000000
D:"Eject Page Interval"=00000000
S:"Monitor Listen Address"=0.0.0.0:22
D:"Monitor Allow Remote Input"=00000000
D:"Disable Resize"=00000002
D:"Auto Reconnect"=00000002
B:"Page Margins"=00000020
 00 00 00 00 00 00 e0 3f 00 00 00 00 00 00 e0 3f 00 00 00 00 00 00 e0 3f 00 00 00 00 00 00 e0 3f
B:"Printer Font v2"=00000060
 f3 ff ff ff 00 00 00 00 00 00 00 00 00 00 00 00 90 01 00 00 00 00 00 00 03 02 01 31 43 00 6f 00
 75 00 72 00 69 00 65 00 72 00 20 00 4e 00 65 00 77 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00
 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 64 00 00 00
D:"Page Orientation"=00000001
D:"Paper Size"=00000001
D:"Paper Source"=00000007
D:"Printer Quality"=fffffffd
D:"Printer Color"=00000001
D:"Printer Duplex"=00000001
D:"Printer Media Type"=00000001
S:"Printer Name"=
D:"Disable Pass Through Printing"=00000000
D:"Buffer Pass Through Printing"=00000000
D:"Force Black On White"=00000000
D:"Use Raw Mode"=00000000
D:"Printer Baud Rate"=00009600
D:"Printer Parity"=00000000
D:"Printer Stop Bits"=00000000
D:"Printer Data Bits"=00000008
D:"Printer DSR Flow"=00000000
D:"Printer DTR Flow Control"=00000001
D:"Printer CTS Flow"=00000001
D:"Printer RTS Flow Control"=00000002
D:"Printer XON Flow"=00000000
S:"Printer Port"=
S:"Printer Name Of Pipe"=
D:"Use Printer Port"=00000000
D:"Use Global Print Settings"=00000001
D:"Operating System"=00000000
S:"Time Zone"=
S:"Last Directory"=
S:"Initial Local Directory V2"=
S:"Default Download Directory V2"=
D:"File System Case"=00000000
S:"File Creation Mask"=
D:"Disable Directory Tree Detection"=00000000
D:"Verify Retrieve File Status"=00000001
D:"Resolve Symbolic Links"=00000002
B:"RemoteFrame Window Placement"=0000002c
 2c 00 00 00 00 00 00 00 01 00 00 00 00 00 00 00 00 00 00 00 fc ff ff ff fc ff ff ff 00 00 00 00
 00 00 00 00 00 00 00 00 00 00 00 00
S:"Remote ExplorerFrame State"=1,1000,200
S:"Remote ListView State"=1,1,1,0,0
S:"SecureFX Remote Tab State"=1,-1,-1
D:"Restart Data Size"=00000000
S:"Restart Datafile Path"=
D:"Max Transfer Buffers"=00000004
D:"Filenames Always Use UTF8"=00000000
D:"Use A Separate Transport For Every Connection"=00000000
D:"Use Multiple SFTP Channels"=00000000
D:"Disable STAT For SFTP Directory Validation"=00000000
D:"Use STAT For SFTP Directory Validation"=00000000
D:"Disable MLSX"=00000000
D:"SecureFX Trace Level V2"=00000002
D:"Synchronize App Trace Level"=00000001
D:"SecureFX Use Control Address For Data Connections"=00000001
D:"Use PGP For All Transfers"=00000000
D:"Disable Remote File System Watches"=00000000
Z:"Port Forward Table V2"=00000002
 Intranet|127.0.0.1,3128|1|webproxy.ncbi.nlm.nih.gov|3128||
 MsLogin01|127.0.0.2,3390|1|MSLOGIN01|3389|C:\windows\system32\mstsc.exe|/v:127.0.0.2:3390
Z:"Reverse Forward Table V2"=00000000
Z:"Keymap v4"=00000000
Z:"Description"=00000000
Z:"SecureFX Post Login User Commands"=00000000
Z:"SecureFX Bookmarks"=00000000
Z:"SCP Shell Prompts"=00000001
 "? ",0,"\n"
'@

  return $DefaultIni
}

function Set-RDP(){
    Try{
    $Test = Test-Path -Path "$env:userprofile\Documents\Default.rdp"
    If (-not $Test){
        $TRAP = "" | Out-File -filepath "$env:userprofile\Documents\Default.rdp" -Force -NoNewline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
    $FILERDP = (GCI -Path "$env:userprofile\Documents\Default.rdp" -Force).FullName
    [array]$Content=$ContentOrg = Get-Content -Path $FILERDP | Sort-Object
    $Content += "enablecredsspsupport:i:0"
    $Content += "authentication level:i:2"
    If (($Content -match "enablecredsspsupport:i:[1-9]")){
            $Content = $Content -replace "enablecredsspsupport:i:.", "enablecredsspsupport:i:0"
            $Changed = $true 
        }
    If (($Content -match "authentication level:i:[13-9]")){
            $Content = $Content -replace "authentication level:i:.", "authentication level:i:2"
            $Changed = $true 
        }
    $Content = $Content | select-string -Pattern "^\w+.*$" | ? {([string]$_.line).Trim() -ne '' } | Foreach {$_ -replace "`n`r",''} | Select-Object -Unique | Sort-Object
    #write-host $Content
    $Trap = Set-ItemProperty $filerdp -name Attributes -Value "Normal"
    $Trap = Set-Content -Path $FILERDP -Force -value $Content
    $Trap = Set-ItemProperty $filerdp -name Attributes -Value "Hidden"
    }
    Catch{
        Throw
    }

}
function Show-Inputbox {
   Param([string]$message=$(Throw "You must enter a prompt message"),
         [string]$title="Input",
         [string]$default
         )
         
         [reflection.assembly]::loadwithpartialname("microsoft.visualbasic") | Out-Null
         [microsoft.visualbasic.interaction]::InputBox($message,$title,$default)
}
Function Show-Messagebox{ 
    [CmdletBinding()][OutputType([int])]
        Param( 
        [parameter(Mandatory=$true, ValueFromPipeLine=$false)][Alias("Msg")][string]$Message, 
        [parameter(Mandatory=$true, ValueFromPipeLine=$false)][Alias("Ttl")][string]$Title = $null, 
        [parameter(Mandatory=$true, ValueFromPipeLine=$false)][Alias("Duration")][int]$TimeOut = 0, 
        [parameter(Mandatory=$true, ValueFromPipeLine=$false)][Alias("But","BS","Button")][ValidateSet( "OK", "OC", "AIR", "YNC" , "YN" , "RC")][string]$ButtonSet = "OK", 
        [parameter(Mandatory=$false, ValueFromPipeLine=$false)][Alias("ICO")][ValidateSet( "None", "Critical", "Question", "Exclamation" , "Information" )][string]$IconType = "None",
        [parameter(Mandatory=$false, ValueFromPipeLine=$false)][switch]$ISSilent = $silent  
         ) 
 
    $ButtonSets = "OK", "OC", "AIR", "YNC" , "YN" , "RC" 
    $IconTypes  = "none", "critical", "question", "exclamation" , "information" 
    $IconVals = 0,16,32,48,64 
    if((Get-Host).Version.Major -ge 3){ 
        $Button   = $ButtonSets.IndexOf($ButtonSet.ToUpper()) 
        $Icon     = $IconVals[$IconTypes.IndexOf($IconType.ToLower())] 
        } 
    else{ 
        $ButtonSets|ForEach-Object -Begin{$Button = 0;$idx=0} -Process{ if($_.Equals($ButtonSet)){$Button = $idx           };$idx++ } 
        $IconTypes |ForEach-Object -Begin{$Icon   = 0;$idx=0} -Process{ if($_.Equals($IconType) ){$Icon   = $IconVals[$idx]};$idx++ } 
        } 
     if (-not $Silent){   
     $window = new-object -comobject wscript.shell
     $return = $window.popup($message,$time,$title,$Button+$Icon) 
     return $return
     }Else{
     Return -1
     }  
}
function Uninstall-MSI{
    [CmdletBinding()]
    [OutputType([boolean])]
    param 
    (
    [string]$program = 'SecureCRT',
    [Version]$version = '8.5.0',
    [boolean]$uninstall=$true
    )
    $path = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    # get all data
    
    $RS = Get-ItemProperty $path | Where-Object { ($_.DisplayName -match $program -or $_.DisplayName -match 'securefx')  -and [version]$_.displayversion -lt [version]$version } | Sort-Object -Property DisplayName
    foreach ($R in $RS){
        $stopproc = Get-Process | ? {$_.name -match 'secure' -or $_.name -match 'slack'}| Stop-Process -Force
        $String = $R.PSChildName
        if ($uninstall){
            $Params = '/x',$string,'/qb-','/norestart' 
            $p = Start-Process 'msiexec.exe' -ArgumentList $params -PassThru -Verb RunAs
            $count= 0
            while ($false -eq $p.HasExited){
                Start-Sleep -Seconds 3
                If ($count -ge 100){
                    Break
                }
                Else{
                $Count++
                }
            }
            
        }
    }
    Return $true
}
Function Validate-PreinstallEnvironmentCheck {
  <#
      .SYNOPSIS
      xxxxxxxxxxxxxxxxxx
      .DESCRIPTION
      xxxxxxxxxxxxxxxxxx
      .PARAMETER InputObject
      xxxxxxxxxxxxxxxxxx
      .PARAMETER PropertyName
      xxxxxxxxxxxxxxxxxx
      .PARAMETER ArgumentList
      xxxxxxxxxxxxxxxxxx
      .EXAMPLE
      xxxxxxxxxxxxxxxxxx
      .NOTES
      xxxxxxxxxxxxxxxxxx
      .LINK
      xxxxxxxxxxxxxxxxxx
  #>
  [CmdletBinding()]

  Param (
    [Parameter(Mandatory=$false,Position=0)]
    [ValidateNotNull()]
    [string]$PathFunc,
    [Parameter(Mandatory=$false,Position=1)]
    [ValidateNotNullorEmpty()]
    [string]$PropertyName,
    [Parameter(Mandatory=$false,Position=2)]
    [object[]]$ArgumentList
  )
	
  #Check Admin
    $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
    $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
    $IsAdmin=$prp.IsInRole($adm)
    #Check Architecture
    $pcArch = ""
    if (($ENV:Processor_Architecture -eq "x86" -and (test-path env:PROCESSOR_ARCHITEW6432)) -or ($ENV:Processor_Architecture -eq "AMD64")) {
        $pcArch = "x64"
        Write-Debug "PC Architecture is $pcArch"
    } 
    #CheckOS
    [psobject]$envOS = Get-WmiObject -Class 'Win32_OperatingSystem' -ErrorAction 'SilentlyContinue'
    [string]$envOSName = $envOS.Caption.Trim()
    [boolean]$Win10AtLeast = [version]$envOS.Version -ge [version]"10.0.0" 
    if ($Win10AtLeast) {
        $WinVer= "$envosname"
        Write-Debug "$envosname"
    }  
    #Check Files
    $MSIExist = GCI -path "$pathfunc\*" -Include '*.msi' 
    IF ($MSIExist){
        $MSIExist= $MSIExist
        Write-Debug $MSIExist
    }  
    IF ($MSIExist -and $IsAdmin -and $envOSName){
        return $true
    }
    Else{
    $Message = "One of the prerequistises did not pass. `nFile Exists in Files Subdirectory `n$MSIExist `nIs a local adminsitrator `n$Isadmin `nIs at least Windows 10 `n$envosname"
    $Title = "Cannot Continue"
    $timeout = 400
    $Buttonset ='ok'
    $icontype ='critical'
    $TRAP =Show-Messagebox -Message $message -Title $Title -TimeOut $timeout  -ButtonSet $buttonset -icontype $icontype 
    Break
    }

}
##*************************
##*************************
## Variables: Environment
$ConfirmPreference = 'high'
$WarningPreference = 'silentlycontinue'
$ErrorActionPreference='SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$VerbosePreference = 'Silentlycontinue'
$start =Start-Transcript -Path $env:TEMP\SecCRTProfile.log -Append -Force
If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation
  [string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
 }ElseIf ($MyInvocation.MyCommand.CommandType -eq "ExternalScript"){
  $ScriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
 }ElseIf ($psscriptroot){
  $ScriptDirectory = $psscriptroot
 }
#Get-Variable * | Out-Host
$files = "$scriptDirectory\files"
#write-host $files
## Initialize
#Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

$CertsSmartCard=$CertsSmartCardNot=$CertstokeepSAN=$CertstokeepSANNot=$CertstokeepSANOverflow=$CertsExpired=$CertsExpiredNot=$CertsAffialiteA=$CertsAffialiteANot = New-Object psobject | Add-Member NoteProperty -name Thumbprint -Value "00 00 00 00"
$user = $env:USERNAME
$domain= $env:USERDOMAIN
$ekuName = "Smart Card Logon" # '-a credential'
$sccert = ""
$lastrun = ""
$smtpserver = ""
$mailfrom = ""
$mailto = ""
$PIVmatchAdCertbool=$false
# Where SecureCRT looks to find sessions
$configSessionPath = $env:APPDATA + "\VanDyke\Config\Sessions"
$configSession = $env:APPDATA + "\VanDyke"
#The name of the session file to be created
$sshSessionName = "NCBI PIV Remote Access.ini"
## Initialize Functions
$UserCertA = ""
$CleanCerts = Clean-CerticateStore #Removes Expired -A certs
$CertsSmartCard,$CertsSmartCardNot,$CertstokeepSAN,$CertstokeepSANNot,$CertstokeepSANOverflow,$CertsExpired,$CertsExpiredNot,$CertsAffialiteA,$CertsAffialiteANot = Get-UserCertificates
##Execute Code
$Validate = Validate-PreinstallEnvironmentCheck -PathFunc $files
##Uninstall Previous
$Files = gci -Path $files -include '*.msi' -Recurse -File | Select-Object -First 1 #only first msi
$SecureCRTInstall = get-msiproperties -MsiPath $files	
if ($SecureCRTInstall.productversion){
  $Software = Uninstall-MSI -program 'securecrt' -version ($SecureCRTInstall.productversion) -uninstall $true
}
$Install = Install-Application
$CopyLicense = Copy-License
$CopyINI = Copy-Ini
$SETRDP= Set-RDP
$Message = "Install is complete."
$Title = "Successful Install"
$timeout = 400
$Buttonset ='ok'
$icontype ='INFORMATION'
$VAL= Show-Messagebox -Message $message -Title $Title -TimeOut $timeout  -ButtonSet $buttonset -icontype $icontype 
$start =Stop-Transcript
