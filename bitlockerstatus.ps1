function Get-MBAMClientDetails {
# version 1.1
[CmdletBinding()]

Param (
    [parameter(ValueFromPipeline=$True)]
    [string[]]$ComputerName
    )

Begin {
    #Initialize
    Write-Verbose "Initializing"
    if($ComputerName -eq ".") {
        $ComputerName = $env:COMPUTERNAME
    }
    if ($ComputerName -eq $null ) {
        $ComputerName = $env:COMPUTERNAME
    }
}
    

Process {
    #---------------------------------------------------------------------
    # Process each ComputerName
    #---------------------------------------------------------------------
    if (!($PSCmdlet.MyInvocation.BoundParameters[“Verbose”].IsPresent)) {
        Write-Host "Processing $ComputerName"
        }

    Write-Verbose "Processing $ComputerName "

    $htmlreport = @()
    $htmlbody = @()
    $htmlfile = "$($ComputerName).html"
    $spacer = "<br />"

    #---------------------------------------------------------------------
    # Do 10 pings and calculate the fastest response time
    # Not using the response time in the report yet so it might be
    # removed later.
    #---------------------------------------------------------------------
    
    try {
        $bestping = (Test-Connection -ComputerName $ComputerName -Count 10 -ErrorAction STOP | Sort ResponseTime)[0].ResponseTime
        }
        catch {
            Write-Warning $_.Exception.Message
            $bestping = "Unable to connect"
            }

    if ($bestping -eq "Unable to connect") {
        if (!($PSCmdlet.MyInvocation.BoundParameters[“Verbose”].IsPresent)) {
            Write-Host "Unable to connect to $ComputerName"
            }
        "Unable to connect to $ComputerName"
        }
        else {

            #---------------------------------------------------------------------
            # Collect computer system information and convert to HTML fragment
            #---------------------------------------------------------------------    
            Write-Verbose "Collecting Computer System Information"
            $subhead = "<h3>Computer System Information</h3>"
            $htmlbody += $subhead
            Try {
            $csinfo = Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName -ErrorAction STOP |
                Select Name,Manufacturer,Model,
                            @{Name='Physical Processors';Expression={$_.NumberOfProcessors}},
                            @{Name='Logical Processors';Expression={$_.NumberOfLogicalProcessors}},
                            @{Name='Total Physical Memory (Gb)';Expression={
                                $tpm = $_.TotalPhysicalMemory/1GB;
                                "{0:F0}" -f $tpm
                            }},
                            DnsHostName,Domain
       
            $htmlbody += $csinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
       
            }
            Catch {
                Write-Warning $_.Exception.Message
                $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
                $htmlbody += $spacer
                }
        
#---------------------------------------------------------------------
# Collect OSD Build Information and convert to HTML fragment
#---------------------------------------------------------------------
         Write-Verbose "Collecting OSD Build Information"

        $subhead = "<h3>OSD Build Information</h3>"
        $htmlbody += $subhead

        Try {
            $OSDBuild = Invoke-Command -ComputerName $ComputerName -ErrorAction STOP -ScriptBlock {
                $TSOSD = (Get-ItemProperty –Path “HKLM:\Software\Microsoft\Deployment 4”)
                $TSOSD | Select "THD Site Number","Task Sequence Name","THD Lifecycle"
                }
            }
            Catch {
                Write-Warning $_.Exception.Message
                $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
                $htmlbody += $spacer
                }        

            $htmlbody += $OSDBuild | Select "THD Site Number","Task Sequence Name","THD Lifecycle" | ConvertTo-Html -Fragment
            $htmlbody += $spacer
                
#---------------------------------------------------------------------
# Collect operating system information and convert to HTML fragment
#---------------------------------------------------------------------
    
        Write-Verbose "Collecting Operating System Information"

        $subhead = "<h3>Operating System Information</h3>"
        $htmlbody += $subhead
    
        Try {
            $osinfo = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction STOP | 
            Select @{Name='Operating System';Expression={$_.Caption}},
                   @{Name='Architecture';Expression={$_.OSArchitecture}},
                   Version,Organization,
                   @{Name='Install Date';Expression={
                   $installdate = [datetime]::ParseExact($_.InstallDate.SubString(0,8),"yyyyMMdd",$null);
                   $installdate.ToShortDateString()
                            }}

            $htmlbody += $osinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        Catch {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            }

#---------------------------------------------------------------------
# Collect physical memory information and convert to HTML fragment
#---------------------------------------------------------------------

        Write-Verbose "Collecting physical memory information"

        $subhead = "<h3>Physical Memory Information</h3>"
        $htmlbody += $subhead

        Try {
            $memorybanks = @()
            $physicalmemoryinfo = @(Get-WmiObject Win32_PhysicalMemory -ComputerName $ComputerName -ErrorAction STOP |
                Select DeviceLocator,Manufacturer,Speed,Capacity)

            foreach ($bank in $physicalmemoryinfo)
            {
                $memObject = New-Object PSObject
                $memObject | Add-Member NoteProperty -Name "Device Locator" -Value $bank.DeviceLocator
                $memObject | Add-Member NoteProperty -Name "Manufacturer" -Value $bank.Manufacturer
                $memObject | Add-Member NoteProperty -Name "Speed" -Value $bank.Speed
                $memObject | Add-Member NoteProperty -Name "Capacity (GB)" -Value ("{0:F0}" -f $bank.Capacity/1GB)

                $memorybanks += $memObject
            }

            $htmlbody += $memorybanks | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        Catch {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            }

#---------------------------------------------------------------------
# Collect BIOS information and convert to HTML fragment
#---------------------------------------------------------------------

        $subhead = "<h3>BIOS Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting BIOS information"

        Try {
            $biosinfo = Get-WmiObject Win32_Bios -ComputerName $ComputerName -ErrorAction STOP |
                Select Status,Version,Manufacturer,SMBIOSBIOSVersion,
                @{Name='Release Date';Expression={
                    $releasedate = [datetime]::ParseExact($_.ReleaseDate.SubString(0,8),"yyyyMMdd",$null);
                    $releasedate.ToShortDateString()
                }}

            $htmlbody += $biosinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
            }
            Catch {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            }

#---------------------------------------------------------------------
# Collect logical disk information and convert to HTML fragment
#---------------------------------------------------------------------

        $subhead = "<h3>Logical Disk Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting logical disk information"

        Try {
            $diskinfo = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -ErrorAction STOP | 
            Select DeviceID,FileSystem,VolumeName,
            @{Expression={$_.Size /1Gb -as [int]};Label="Total Size (GB)"},
            @{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Space (GB)"}

            $htmlbody += $diskinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
            }
            Catch {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            }

#---------------------------------------------------------------------
# Collect BitLocker Drive Encryption Information and convert to HTML fragment
#---------------------------------------------------------------------
<#
ComputerName: HD6553EF37D5EA3

VolumeType      Mount CapacityGB VolumeStatus           Encryption KeyProtector              AutoUnlock Protection
                Point                                   Percentage                           Enabled    Status    
----------      ----- ---------- ------------           ---------- ------------              ---------- ----------
Data            \\...       0.49 FullyDecrypted         0          {}                                   Off       
OperatingSystem C:        236.88 FullyDecrypted         0          {}                                   Off       
#>
        $subhead = "<h3>BitLocker Drive Encryption Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting BitLocker Drive Encryption Information"

        Try {
            $BitLocker = Get-WmiObject -ComputerName $ComputerName -Class Win32_EncryptableVolume -Namespace root\CIMV2\Security\MicrosoftVolumeEncryption -ErrorAction STOP     
            #$BitLocker = Get-WmiObject -ComputerName $ComputerName -Class Win32_EncryptableVolume -Namespace root\CIMV2\Security\MicrosoftVolumeEncryption -Filter "DriveLetter = 'c:'" -ErrorAction STOP     
            #$BitLocker = Get-WmiObject -Class Win32_EncryptableVolume -Namespace root\CIMV2\Security\MicrosoftVolumeEncryption -Filter "DriveLetter = 'c:'" -ErrorAction STOP    
            # Bitlocker Volume Type

            #$Bitlocker.DriveLetter
            
            $VolType = Switch($BitLocker.VolumeType) {
            0 {"OSVolume"}
            1 {"FixedDataVolume"}
            2 {"PortableDataVolume"}
            }
            # Protection Status
            $EncryptVol = $BitLocker.ProtectionStatus
            $ProStat = switch ($EncryptVol) {
            0 {"Protection OFF"}
            1 {"Protection ON (Unlocked)"}
            2 {"Protection ON (Locked)"}
            }
            # Get PCR[7] Secure Boot Binding State
            $BindingState = Switch ($BitLocker.GetSecureBootBindingState().BindingState) {
            0 {"Not Possible"}
            1 {"Disabled By Policy"}
            2 {"Possible"}
            3 {"Bound"}
            }
            # Lock Status
            $LockStatus = Switch ($BitLocker.GetLockStatus().LockStatus) {
            0 {"Unlocked"}
            1 {"Locked"}
            }
            # Encryption Method
            $Cipher = $BitLocker.GetEncryptionMethod().encryptionmethod
            $EncryptMethod = switch ($Cipher){
            '0'{'None';break}
            '1'{'AES_128_WITH_DIFFUSER';break}
            '2'{'AES_256_WITH_DIFFUSER';break}
            '3'{'AES_128';break}
            '4'{'AES_256';break}
            '5'{'HARDWARE_ENCRYPTION';break}
            '6'{'XTS-AES 128';break}
            }
            # AutoUnlockEnabled - The IsAutoUnlockEnabled method of the Win32_EncryptableVolume class indicates whether the volume is automatically unlocked when it is mounted (for example, when removable memory devices are connected to the computer). 
            # Note: Note Applicable to operating system drive
            $IsAutoUnlockEnabled = $BitLocker.IsAutoUnlockEnabled().IsAutoUnlockEnabled
            # AutoUnlockKeyStore
            $IsAutoUnlockKeyStore = ($BitLocker.IsAutoUnlockKeyStored()).IsAutoUnlockKeyStored
            # Volume Status
            $VolumeProtection = switch ($BitLocker.GetConversionStatus().Conversionstatus){
            '0' {'Fully Decrypted'}
            '1' {'Fully Encrypted'}
            '2' {'Encryption In Progress'}
            '3' {'Decryption In Progress'}
            '4' {'Encryption Paused'}
            '5' {'Decryption Paused'}                    
            }
            # Get Key Protector(s) Type
            $ProtectorIds = $BitLocker.GetKeyProtectors("0").volumekeyprotectorID            
            [string[]]$return = @()
            foreach ($ProtectorID in $ProtectorIds){ 
                $KeyProtectorType = $BitLocker.GetKeyProtectorType("$ProtectorID").KeyProtectorType
                switch($KeyProtectorType){
                    "0"{$return += "Unknown or other protector type";break}
                    "1"{$return += "Trusted Platform Module (TPM)";
                        $return += "Platform Validation Profile " + "{ " + $BitLocker.GetKeyProtectorPlatformValidationProfile("$ProtectorID").PlatformValidationProfile + " }" }
                    "2"{$return += "External key";break}
                        # GetKeyProtectorExternalKey(    
                    "3"{$return += "Numerical password";
                        $return += $BitLocker.GetKeyProtectorNumericalPassword("{35A95BD8-56C3-4A4C-A6F0-7CBB37F749D5}").NumericalPassword}
                    "4"{$return += "TPM And PIN";break}
                        # $BitLocker.GetKeyProtectorPlatformValidationProfile("$ProtectorID")
                    "5"{$return += "TPM And Startup Key";break}
                        # $BitLocker.GetKeyProtectorPlatformValidationProfile("$ProtectorID")
                        # GetKeyProtectorExternalKey(       
                    "6"{$return += "TPM And PIN And Startup Key";break}
                        # $BitLocker.GetKeyProtectorPlatformValidationProfile("$ProtectorID")
                        # GetKeyProtectorExternalKey(
                    "7"{$return += "Public Key";break}
                        # GetKeyProtectorCertificate(
                    "8"{$return += "Passphrase";break}
                    "9"{$return += "TPM Certificate";break}
                    "10"{$return += "CryptoAPI Next Generation (CNG) Protector";break}
                    }
                    }

                $BDEStatus = $BitLocker | ForEach {
                    $props = [ordered]@{
                        "Suspend Protection Reboot Count" = $_.GetSuspendCount().SuspendCount
                        "PCR7 Configuration" = $BindingState
                        "Encryption Method" = $EncryptMethod
                        "Volume Status" = $VolumeProtection
                        "Protection Status" = $ProStat                    
                        "Lock Status" = $LockStatus 
                        "Percentage Encrypted" = $_.GetConversionStatus().EncryptionPercentage
                        "Volume Type" = $VolType
                        "Auto Unlock Enabled" = $IsAutoUnlockEnabled
                        "Auto Unlock Key Stored" = $IsAutoUnlockKeyStore
                        "Key Protector(s)" = [pscustomobject]$return
                        }
                    New-Object PsObject -Property $props
                    }            
                
            $htmlbody += $BDEStatus | ConvertTo-Html -Fragment
            $htmlbody += $spacer
            }
            Catch {
                Write-Warning $_.Exception.Message
                $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
                $htmlbody += $spacer
                }

#---------------------------------------------------------------------
# Collect Trusted Platform Module Information and convert to HTML fragment
#---------------------------------------------------------------------

        $subhead = "<h3>TPM Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting Trusted Platform Module Information"

        Try {
            $TPM = GWMI -Namespace ROOT\CIMV2\Security\MicrosoftTpm -Class Win32_Tpm -ComputerName $ComputerName -ErrorAction STOP

            $Information = @{
            #INFORMATION_SHUTDOWN
            0x00000002 = "Platform restart is required (shutdown)."
            #INFORMATION_REBOOT
            0x00000004 = "Platform restart is required (reboot)."
            #INFORMATION_TPM_FORCE_CLEAR
            0x00000008 = "The TPM is already owned. Either the TPM needs to be cleared or the TPM owner authorization value needs to be imported."
            #INFORMATION_PHYSICAL_PRESENCE
            0x00000010 = "Physical Presence is required to provision the TPM."
            #INFORMATION_TPM_ACTIVATE
            0x00000020 = "The TPM is disabled or deactivated."
            #INFORMATION_TPM_TAKE_OWNERSHIP
            0x00000040 = "The TPM ownership was taken."
            #INFORMATION_TPM_CREATE_EK
            0x00000080 = "An Endorsement Key (EK) exists in the TPM."
            #INFORMATION_TPM_OWNERAUTH
            0x00000100 = "The TPM owner authorization is not properly stored in the registry."
            #INFORMATION_TPM_SRK_AUTH
            0x00000200 = "The Storage Root Key (SRK) authorization value is not all zeros."
            #INFORMATION_TPM_DISABLE_OWNER_CLEAR
            0x00000400 = "If the operating system is configured to disable clearing of the TPM with the TPM owner authorization value and the TPM has not yet been configured to prevent clearing of the TPM with the TPM owner authorization value ."
            #INFORMATION_TPM_SRKPUB
            0x00000800 = "The operating system's registry information about the TPM’s Storage Root Key does not match the TPM Storage Root Key."
            #INFORMATION_TPM_READ_SRKPUB
            0x00001000 = "The TPM permanent flag to allow reading of the Storage Root Key public value is not set."
            #INFORMATION_TPM_BOOT_COUNTER
            0x00002000 = "The monotonic counter incremented during boot has not been created."
            #INFORMATION_TPM_AD_BACKUP
            0x00004000 = "The TPM’s owner authorization has not been backed up to Active Directory."
            #INFORMATION_TPM_AD_BACKUP_PHASE_I
            0x00008000 = "The first portion of the TPM owner authorization information storage in Active Directory is in progress."
            #INFORMATION_TPM_AD_BACKUP_PHASE_II
            0x00010000 = "The second portion of the TPM owner authorization information storage in Active Directory is in progress."
            #INFORMATION_LEGACY_CONFIGURATION
            0x00020000 = "Windows Group Policy is configured to not store any TPM owner authorization so the TPM cannot be fully ready."
            #INFORMATION_EK_CERTIFICATE
            0x00040000 = "The EK Certificate was not read from the TPM NV Ram and stored in the registry."
            #INFORMATION_TCG_EVENT_LOG
            0x00080000 = "The TCG event log is empty or cannot be read."
            #INFORMATION_NOT_REDUCED
            0x00100000 = "The TPM is not owned."
            #INFORMATION_GENERIC_ERROR
            0x00200000 = "An error occurred, but not specific to a particular task."
            #INFORMATION_DEVICE_LOCK_COUNTER
            0x00400000 = "The device lock counter has not been created."
            #INFORMATION_DEVICEID
            0x00800000 = "The device identifier has not been created."
            }
            $RdyBitMsk = $TPM.IsReadyInformation().Information
            $ReadyInfo = $Information.Keys | where { $_ -band $RdyBitMsk } | foreach { $Information.Get_Item($_) }

            #Decode ManufacturerID
            $ManufacturerID = $tpm.ManufacturerId
            $HEX = [Convert]::ToString($ManufacturerID, 16)
            $Decoded = [regex]::Matches($HEX,'..')|%{$_.Value}
            $DecodedManufacturerID =(@( $Decoded.split(" ") | FOREACH {  ([CHAR][BYTE]([CONVERT]::toint16($_,16))) })-join '') 

            $TPMStatus = GWMI -ComputerName $ComputerName -Namespace ROOT\CIMV2\Security\MicrosoftTpm -Class Win32_Tpm -ErrorAction STOP | ForEach {
                $props = [ordered]@{
                    "Specification Version"= $_.SpecVersion.Split(',')[0]
                    "Manufacturer Name" = $DecodedManufacturerID
                    "Physical Presence Version" = $_.PhysicalPresenceVersionInfo
                    "TPM Activated" = $_.IsActivated().IsActivated
                    "TPM Enabled" = $_.IsEnabled().IsEnabled
                    "Auto Provision Enabled" = $_.IsAutoProvisioningEnabled().IsAutoProvisioningEnabled
                    "Owner Clear Disabled" = $_.IsOwnerClearDisabled().IsOwnerClearDisabled
                    "Ownership Allowed" = $_.IsOwnershipAllowed().IsOwnershipAllowed
                    "Physical Clear Disabled" = $_.IsPhysicalClearDisabled().IsPhysicalClearDisabled
                    "SRK Auth Compatible" = $_.IsSrkAuthCompatible().IsSrkAuthCompatible
                    "TPM Owned" = $_.IsOwned().IsOwned
                    "TPM Owner Auth" = $_.GetOwnerAuth().OwnerAuth
                    "TPM Endorsement Key Pair Present" = $_.IsEndorsementKeyPairPresent().IsEndorsementKeyPairPresent
                    "Is Ready" = $_.IsReady().IsReady 
                    "Is Ready Information" = [pscustomobject]$ReadyInfo
                }
        
                New-Object PsObject -Property $props

                }

            $htmlbody += $TPMStatus | Select "Specification Version","Manufacturer Name","Physical Presence Version","TPM Activated",
            "TPM Enabled","Auto Provision Enabled","Owner Clear Disabled" | ConvertTo-Html -Fragment
            $htmlbody += $spacer

            $htmlbody += $TPMStatus | Select "Ownership Allowed","Physical Clear Disabled","SRK Auth Compatible","TPM Owned","TPM Owner Auth","TPM Endorsement Key Pair Present" | ConvertTo-Html -Fragment
            $htmlbody += $spacer

            $htmlbody += $TPMStatus | Select "Is Ready","Is Ready Information" | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        
        }
        catch {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            }

#---------------------------------------------------------------------
# Collect network interface information and convert to HTML fragment
#---------------------------------------------------------------------    

        $subhead = "<h3>Network Interface Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting Network Interface Information"
        #-ComputerName $ComputerName
        Try {
            $nics = @()
            $nicinfo = @(Get-WmiObject Win32_NetworkAdapter -ErrorAction STOP | Where {$_.PhysicalAdapter} |
                Select-Object Name,AdapterType,MACAddress,
                @{Name='ConnectionName';Expression={$_.NetConnectionID}},
                @{Name='Enabled';Expression={$_.NetEnabled}},
                @{Name='Speed';Expression={$_.Speed/1000000}})
                #-ComputerName $ComputerName
            $nwinfo = Get-WmiObject Win32_NetworkAdapterConfiguration  -ErrorAction STOP |
                Select-Object Description, DHCPServer,  
                @{Name='IpAddress';Expression={$_.IpAddress -join '; '}},  
                @{Name='IpSubnet';Expression={$_.IpSubnet -join '; '}},  
                @{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}},  
                @{Name='DNSServerSearchOrder';Expression={$_.DNSServerSearchOrder -join '; '}}

            foreach ($nic in $nicinfo)
            {
                $nicObject = New-Object PSObject
                $nicObject | Add-Member NoteProperty -Name "Connection Name" -Value $nic.connectionname
                $nicObject | Add-Member NoteProperty -Name "Adapter Name" -Value $nic.Name
                $nicObject | Add-Member NoteProperty -Name "Type" -Value $nic.AdapterType
                $nicObject | Add-Member NoteProperty -Name "MAC" -Value $nic.MACAddress
                $nicObject | Add-Member NoteProperty -Name "Enabled" -Value $nic.Enabled
                $nicObject | Add-Member NoteProperty -Name "Speed (Mbps)" -Value $nic.Speed
        
                $ipaddress = ($nwinfo | Where {$_.Description -eq $nic.Name}).IpAddress
                $nicObject | Add-Member NoteProperty -Name "IPAddress" -Value $ipaddress

                $nics += $nicObject
            }

            $htmlbody += $nics | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            }

#---------------------------------------------------------------------
# Collect software information and convert to HTML fragment
#---------------------------------------------------------------------

        $subhead = "<h3>Software Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting Software Information"
        
        Try {
            $software = Get-CimInstance -ComputerName $ComputerName -Query "SELECT * FROM SMS_InstalledSoftware" -Namespace "root\CIMV2\sms" | Select Publisher,ProductName,ProductVersion, InstallDate
                        # Get-WmiObject Win32reg_addremoveprograms -ComputerName $ComputerName -ErrorAction STOP | Select-Object Publisher,DisplayName,Version | Sort-Object Publisher,DisplayName
                        # Get-WmiObject -ComputerName $ComputerName -query "SELECT * FROM SMS_InstalledSoftware" -namespace "root\CIMV2\sms" | Select ProductName,ProductVersion,@{N='Installed On';E={$_.ConvertToDateTime($_.InstallDate).ToString("MM/dd/yyyy")}}
            $htmlbody += $software | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        Catch {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            }
       
#---------------------------------------------------------------------
# Generate the HTML report and output to file
#---------------------------------------------------------------------
	
        Write-Verbose "Producing HTML report"
    
        $reportime = Get-Date

        #Common HTML head and styles
	    $htmlhead="<html>
				    <style>
				    BODY{font-family: Arial; font-size: 8pt;}
				    H1{font-size: 20px;}
				    H2{font-size: 18px;}
				    H3{font-size: 16px;}
                    TABLE{
                        border-collapse: collapse;
                        border-bottom: 1px solid #CECECE; // gray for last line of table border
                        font-size: 8pt;
                    }
                    TH{
                        border-bottom: 1px solid black; 
                        //background: #dddddd;
                        text-align: left; 
                        padding: 5px; 
                        color: #000000;}
				    TD{                        
                        border-bottom: 1px solid black 
                        padding: 5px; 
                    }
				    td.pass{background: #7FFF00;}
				    td.warn{background: #FFE600;}
				    td.fail{background: #FF0000; color: #ffffff;}
				    td.info{background: #85D4FF;}
                    tr:hover {background-color: coral;}
				    </style>
				    <body>
				    <h1 align=""center"">Detailed Report for: $($ComputerName.ToUpper())</h1>
				    <h3 align=""center"">Generated: $reportime</h3>"

        $htmltail = "</body>
			    </html>"

        $htmlreport = $htmlhead + $htmlbody + $htmltail

        $htmlreport | Out-File $env:TEMP\$htmlfile -Encoding Utf8
        Write-Host "File is saved at $env:temp\$htmlfile" -ForegroundColor Yellow
        Invoke-Item $env:temp\$htmlfile
    }

}

End
{
    #Wrap it up
    Write-Verbose "=====> Finished <====="
}
}