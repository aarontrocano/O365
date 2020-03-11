<#
    .SYNOPSIS
    This tool changes the local WINRM registry settings that are required to connect to Office365 with Basic authentication and Multifactor Authentication
    This tool checks the local WINRM registry settings
    This tool reverts the local WINRM registry settings
    This tool connects to Office365 with Basic authentication
    This tool connects to Office365 with Multifactor Authentication
    This tool removes the Active PSsession to Office365
    on
    .NOTES
    This script was written to get over the hurtle of connecting to Office365 with different authentication types by simplifying the process in an easy to use tool. The script is particular useful in simplifying the 
    MFA process of connecting to OFfice365 by copying the Microsoft Exchange Online Powershell Module to the default Modules folder for the user if it doesn't exist and importing that module into the currently running PSSession.
    This allows use of the MFA Module without having to open the separate Module GUI. Note:  

    #Written By: Ryan T Adcox 12/2019
    #>
function Connect-O365 {
    function Get-MenuOptions {
        param (
            [string]$Title = 'O365 Connectivity Tool'
        )
        Clear-Host
        Write-Host "================ $Title ================"   
        Write-Host "1: Press '1' to Check Registry settings to allow O365 Connectivity."
        Write-Host "2: Press '2' to Change Registry settings to allow O365 Connectivity."
        Write-Host "3: Press '3' to Revert Registry settings to allow O365 Connectivity."
        Write-host "4: Press '4' to Connect to O365 with Basic Authentication!"
        Write-host "5: Press '5' to Connect to O365 with Multi-Factor Authentication!"
        Write-host "6: Press '6' to Remove Active O365 PSSession (NOTE: This is always a good idea once you are done with your current PSSession!)"
        Write-Host "7: Press 'Q' to Quit out of the Menu."
    }
    #Function to Reset PSProfile, this is to reload registry settings in the PSSession without exiting the initial PSsession
    function Reset-Profile {
        @(
            $Profile.AllUsersAllHosts,
            $Profile.AllUsersCurrentHost,
            $Profile.CurrentUserAllHosts,
            $Profile.CurrentUserCurrentHost
        ) | foreach-object {
            if (Test-Path $_) {                       
                Write-Verbose "Running $_" -verbose
                . $_
            }
        }    
    }
    $S = $([Environment]::NewLine)
    do {
        Get-MenuOptions
        $input = Read-Host "Please make a selection"
        switch ($input) {
            '1' {
                Clear-host
                $registrypath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client'
                $winrmquery = get-itemproperty -path "$registrypath"
                $props = [pscustomobject]@{

                    RegistryPath              = $Registrypath
                    Allow_BasicAtuhentication = $winrmquery.allowbasic
                    Allow_UnencryptedTraffic  = $winrmquery.allowunencryptedtraffic
                        
                }
                $props | select-object RegistryPath, Allow_BasicAtuhentication, Allow_UnencryptedTraffic | FL
                #Report if WINRM settings are preventing authentication to Office365
                if ($winrmquery.allowbasic -eq 0 -and $winrmquery.allowunencryptedtraffic -eq 0) {
                    write-warning "This System is preventing Basic Authentication and Unencrypted Traffic via WinRM, this is most likely do to a local or GPO setting to: '$registrypath'" -verbose
                }
                elseif ($winrmquery.allowbasic -eq 1 -and $winrmquery.allowunencryptedtraffic -eq 1) {
                    write-verbose "This System has the required WinRM settings to connect to O365 via Basic Authentication or MFA!$S" -verbose               
                }                         
            } '2' {
                Clear-Host
                $registrypath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client'
                $Name1 = "AllowBasic"    
                $Name2 = "AllowUnencryptedTraffic"
                $Value = "1"                    
                $winrmquery = get-itemproperty -path "$registrypath"
                $test = test-path "$registrypath"
                if ($test -eq $true -and $winrmquery.AllowBasic -eq 0 -and $winrmquery.AllowUnencryptedTraffic -eq 0) {                                  
                    Set-Itemproperty -path "$registrypath" -Name $Name1 -value $Value 
                    Set-Itemproperty -path "$registrypath" -Name $Name2 -value $Value
                    write-verbose "The Registry settings were successfully Changed!" -Verbose                                                                           
                }
                else {
                    Write-Warning "Could not set the Registry Settings at $registrypath!$S" -Verbose
                }
                Reset-Profile                          
            } '3' {
                Clear-Host
                $registrypath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client'
                $Name1 = "AllowBasic"    
                $Name2 = "AllowUnencryptedTraffic"
                $Value = "0"                    
                $winrmquery = get-itemproperty -path "$registrypath"
                $test = test-path "$registrypath"
                if ($test -eq $true -and $winrmquery.AllowBasic -eq 1 -and $winrmquery.AllowUnencryptedTraffic -eq 1) { 
                    write-verbose "The Registry Settings were reverted at $registrypath$S" -Verbose                  
                    Set-Itemproperty -path "$registrypath" -Name $Name1 -value $Value                    
                    Set-Itemproperty -path "$registrypath" -Name $Name2 -value $Value 
                }
                else {
                    Write-Warning "Could NOT revert the Registry Settings at $registrypath!$S" -Verbose
                }                          
            } '4' {
                Clear-Host
                $cred = Get-Credential             
                Import-module MSOnline              
                Connect-MsolService -Credential $Cred            
                $msoExchangeURL = "https://outlook.office365.com/powershell-liveid/"  
                #Connect via Basic Authentication           
                $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $msoExchangeURL -Credential $cred -Authentication Basic -AllowRedirection                
                Import-PSSession $session -Prefix O365
                write-warning "All O365 Commandlets will have a Prefix of 'O365' to segregate them from the Local Exchange Environment ...$S" -Verbose
                write-verbose "Type 'Get-O365mailbox -resultsize 10' to verify that you are connected to Office365 or type 'connect-o365' to bring up the Menu!$S" -verbose 
                Return                          
            } '5' {
                Clear-Host
                #Get Micrsofot Location for Microsoft Exchange Powershell Module for MFA
                $EXOPSSessionlocation = (Get-ChildItem -Path $Env:LOCALAPPDATA\Apps\2.0* -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName
                #Robocopy Options
                $options = @("/e", "/xj", "/r:2", "/w:5", "/v", "/MT:32", "/purge", "/it", "/copy:DAT")
                $source = $EXOPSSessionlocation               
                $destination = ($env:PSModulePath -split ';')[0]
                $Moduledir = ($destination + '\' + 'CreateExoPSSession')           
                $test = Test-Path $Moduledir
                if ($test -eq $false) {
                    #Create New folder in Module directory
                    new-item -ItemType directory -path $Moduledir
                    #Robocopy Module Directory with Options
                    $cmdargs = @($source, $Moduledir, $options)
                    robocopy @cmdargs  
                    #Modifying the CreateExoPSSession.ps1 script to support an O365 Prefix
                    $Stringquery = (get-content -path "$Moduledir\createexopssession.ps1" | select-string -SimpleMatch '$PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber').ToString().Trim() 
                    Write-Verbose "The current String Value in the CreateExoPSSession.ps1 file is:'$stringquery'$S" -Verbose
                    #Replace string in createexopssession.ps1 with string containing the O365 Prefix
                    (get-content -path "$Moduledir\createexopssession.ps1").Replace('$PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber', '$PSSessionModuleInfo = Import-PSSession $PSSession -Prefix O365 -AllowClobber') | set-content -path "$moduledir\createexopssession.ps1"
                    $Stringset = (get-content -path "$Moduledir\createexopssession.ps1" | select-string -SimpleMatch '$PSSessionModuleInfo = Import-PSSession $PSSession -Prefix O365 -AllowClobber').ToString().Trim()
                    Write-Verbose "The String Value in the CreateExoPSSession.ps1 file is now:'$stringset'$S" -Verbose
                    #Importing Module from the new Module Location in Default User Modules
                    Import-module 'createexopssession.ps1' -force -Verbose
                    #Prompt User for UPN
                    $UPN = Read-host -Prompt 'Input Your Office365 UPN (ex: user@example.com)'
                    #Connect via MFA
                    Connect-EXopssession -connectionuri 'https://outlook.office365.com/powershell-liveid/' -userprincipalname $UPN 
                    write-warning "All O365 Commandlets will have a Prefix of 'O365' to segregate them from the Local Exchange Environment ...$S" -Verbose
                    write-verbose "Type 'Get-O365mailbox -resultsize 10' to verify that you are connected to Office365 or type 'connect-o365' to bring up the Menu!$S" -verbose
                    Return             
                }
                else {
                    #Connect to O365 with MFA
                    Import-module 'createexopssession.ps1' -force -Verbose
                    $UPN = Read-host -Prompt 'Input Your Office365 UPN (ex: John.Smith@amtrustgroup.com)'
                    Connect-EXopssession -connectionuri 'https://outlook.office365.com/powershell-liveid/' -userprincipalname $UPN
                    write-verbose "Type 'Get-O365mailbox -resultsize 10' to verify that you are connected to Office365 or type 'connect-o365' to bring up the Menu!$S" -verbose
                    Return
                }                                             
            } '6' {
                try {
                    #Remove Active PSSession
                    Write-verbose "Removed Active PSSession with Office365!$S" -verbose   
                    Get-PSsession | Remove-PSsession 
                }
                catch {
                    Write-Warning "Could NOT Remove Active PSSession with Office365!$S" -Verbose
                }                           
            } 'Q' {
                return
            }
        }
        $error | FL
        pause
    }
    until ($input -eq 'q')
}
#initiate Function
Connect-o365
