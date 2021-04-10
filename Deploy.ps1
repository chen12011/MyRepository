
<#
This script is meant to be run during MDT's task sequence when imaging a new machine. 
This script will install Global Protect, SAP, Microsoft Office, Sophos, and their dependencies.

Executable files are packaged alongside this script for execution.
#>

function Install-GlobalProtect {

        #Initially written as an executable file to display error messages to the user when trying to install global-protect.

        write-host "installing certificates..."

        try{    
            
            #Initializing variables with encrypted password file and certificates

            $eFilePath = $PSScriptRoot + "\encrypt\efile.txt"
            $passFilePath = $PSScriptRoot + "\encrypt\pass.txt"
            $clientCertFilePath = $PSScriptRoot + "\certs\ClientCert.p12"
            $rootCertFilePath = $PSScriptRoot + "\certs\RootCert.p12"
            $sslCertFilePath = $PSScriptRoot + "\certs\SSLCert.crt"


            #Importing pfx certifactes with respectiv encrypted passwords

            Import-PfxCertificate -FilePath $clientCertFilePath -CertStoreLocation Cert:\LocalMachine\My -Password ((Get-Content -Path $passFilePath) | ConvertTo-SecureString -Key (Get-Content -Path $eFilePath)) -ErrorAction Stop  
        
            Import-PfxCertificate -FilePath $rootCertFilePath -CertStoreLocation Cert:\LocalMachine\Root -Password ((Get-Content -Path $passFilePath) | ConvertTo-SecureString -Key (Get-Content -Path $eFilePath)) -ErrorAction Stop  
    
            Import-Certificate -FilePath $sslCertFilePath -CertStoreLocation Cert:\LocalMachine\Root -ErrorAction Stop
        }
        catch{
            $errormessage = $_.exception.message
            
	    $error1= $_.exception
            Write-Host ""
            Write-Host $error1
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            write-host "Certificate installation has failed with the following message:" 
            Write-Host $errormessage 
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            read-host "Press enter to exit"
            break
        }
        Write-Host ""
        write-host "Installing Global Protect..."
        write-host "This may take a few minutes..."
        write-host ""

        try {

            #Installing the Global Protect app using MSI silent switches

            $installArgs = '/i "' + $PSScriptRoot + '\GLobalProtect64.msi" /qb PORTAL="Portal"' #The portal is the name of the Palo Alto Network firewall
            Write-Host "Installing Global Protect..."
            Start-Process "msiexec.exe" -ArgumentList ($installArgs) -Wait
        }   
        catch {

            $errormessage = $_.exception.message

            write-host ""
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            write-host "The installation has failed with the following error:" 
            Write-Host $errormessage 
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            read-host "Press enter to exit"
            break

        }
        write-host ""
        write-host "Changing keys..."
        try{

        #setting up registry keys for always-on VPN
        
        Set-ItemProperty -path "HKLM:\SOFTWARE\Palo Alto Networks\GlobalProtect\PanSetup" -Name "Portal" -Value "Portal" -ErrorAction Stop #The portal is the name of the Palo Alto Network firewall
        Set-ItemProperty -path "HKLM:\SOFTWARE\Palo Alto Networks\GlobalProtect\PanSetup" -Name "prelogon" -Value '1' -ErrorAction Stop
        }
        catch{
            Write-Host ""
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            write-host "Registry key changes has failed with the following error:" 
            Write-Host $errormessage 
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            read-host "Press enter to exit"
            break

        }

        try{
        new-ItemProperty -Path "HKLM:\SOFTWARE\Palo Alto Networks\GlobalProtect\PanSetup" -name "use-sso" -Value "yes" -PropertyType "String" -ErrorAction Stop

        }
        catch{
            write-host ""
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            write-host "Registry key creation has failed with the following error:" 
            Write-Host $errormessage 
            write-host "************************************************************************" -backgroundcolor Red -ForegroundColor White
            read-host "Press enter to exit"
            break
        }


        write-host ""
        write-host "Complete!"
        
    


}

function Install-Office {

    #Installing Microsoft Office 2016 using admin file and appropriate setup files

    Write-Output "`n Installing Microsoft Office..."
    $officePath = $PSScriptRoot + "\Office2016_32bit\setup.exe"
    
    start-process $officePath -ArgumentList "/adminfile Admin_File.MSP" -Wait
    Write-Output "`n Office Install complete!" 
}



function Install-SAP {

    Write-Output "`n Installing SAP GUI..."

    #installing SAP GUI

    $sapPath = $PSScriptRoot + "\SAP_GUI_7.6\BD\PRES1\GUI\WINDOWS\Win32\Setup\NwSapSetup.exe"
    Start-Process -FilePath $sapPath -ArgumentList '/product:"SAPGUI" /noDlg' -Wait
    
    #Copying SAPLogon.ini file to C drive for automated connection setup

    write-output "`n Copying INI file..."
    $iniPath = $PSScriptRoot + "\SAP_GUI_7.6\saplogon.ini"
    Copy-Item $iniPath -Destination "C:\Windows"

    #setting up the SAPLogon.ini as an environmental variable

    Write-Output "`n Creating environmental variable for INI file..."
    [System.Environment]::setEnvironmentVariable('SAPLOGIN_INI_FILE','C:\Windows\saplogon.ini',[System.EnvironmentVariableTarget]::Machine)

    Write-Output "`n SAP installation complete!"
}

function Install-MicrosoftTeams {

    write-output "`n Installing Microsoft Teams..."

    #Installation of Microsoft teams with the use of MSI switches

    $teamsArgs = '/i "' + $PSScriptRoot + '\teams.msi" /qn'
    Start-Process "msiexec.exe" -argumentList $teamsArgs -Wait

    Write-Output "`n Microsoft Teams has completed!"


}

Function Install-AdobeReader {

    Write-Output "`n Installing Adobe Reader...."
 
    #Installation of Adobe-Reader with the use of MSI switches

    $adobeArgs = '/i "' + $PSScriptRoot + '\acroread.msi" /qn'
    Start-Process "msiexec.exe" -argumentList $adobeArgs -Wait

    Write-Output "`n Adobe finished installing!"

}

function Install-Chrome {

    Write-Output "`n Installing Google Chrome..."

    #Installation of Google Chrome with the use of MSI switches

    $chromeArgs = '/i "' + $PSScriptRoot + '\chrome.msi" /qn'
    Start-Process "msiexec.exe" -ArgumentList ($chromeArgs) -Wait

    Write-Output "`n Chrome finished installing!"

}

function Install-Sophos {

    #Installation of Sophos with the use of MSI Switches

    $process = Start-Process "$($PSScriptRoot)\SophosSetup.exe" -ArgumentList "--quiet" -PassThru

    #Script ends when process begins, thus the need to wait till the process disappears before script ends

    $process.WaitForExit()
    get-process | Select-Object Name | Where-Object {$_.Name -match "SophosSetup.exe"} | Wait-Process


}

Install-GlobalProtect
Install-Office
Install-SAP
Install-MicrosoftTeams
Install-AdobeReader
Install-Chrome
Install-Sophos





