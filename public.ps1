#Connect to Exchange Online PowerShell, The first part was taken from https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/. Connecting to exchange through powershell is an ordeal
 Param
 (
    [Parameter(Mandatory = $false)]
    [switch]$Disconnect,
    [switch]$MFA,
    [string]$UserName = "username@domain", 
    [string]$Password = "appPassword"
 )

Set-ExecutionPolicy Bypass -Force -Scope CurrentUser

#Disconnect existing sessions
if($Disconnect.IsPresent)
{
    Get-PSSession | Remove-PSSession
    Write-Host All sessions in the current window has been removed. -ForegroundColor Yellow
}
#Connect Exchnage Online with MFA
elseif($MFA.IsPresent)
{
    #Check for MFA mosule
    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    If ($MFAExchangeModule -eq $null)
    {
        Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow
        Write-Host You can install module using below blog : `nLink `nOR you can install module directly by entering "Y"`n
        $Confirm= Read-Host Are you sure you want to install module directly? [Y] Yes [N] No
        if($Confirm -match "[yY]")
        {
            Write-Host Yes
            Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
        }
        else
        {
            Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
            Exit
        }
        $Confirmation= Read-Host Have you installed Exchange Online MFA Module? [Y] Yes [N] No
        if($Confirmation -match "[yY]")
        {
            $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
            If ($MFAExchangeModule -eq $null)
            {
                Write-Output "Exchange Online MFA module is not available"
                Exit
            }
        }
        else
        { 
            Write-Output "Exchange Online PowerShell Module is required"
            Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
            Exit
        }   
    }
  
    #Importing Exchange MFA Module
    write-host aaaa. "$MFAExchangeModule"
    Connect-EXOPSSession -WarningAction SilentlyContinue | Out-Null
}
#Connect Exchnage Online with Non-MFA
#Connect Exchnage Online with Non-MFA
else
{
    if(($UserName -ne "") -and ($Password -ne "")) 
    { 
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force 
        $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword 
    } 
    else 
    { 
        $Credential=Get-Credential -Credential $null
    } 
  
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null
}

#Check for connectivity
If(!($Disconnect.IsPresent)){
    If ((Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" }) -ne $null)
    {
        Write-Output "Successfully connected to Exchange Online"
    }
    else
    {
        Write-Output "Unable to connect to Exchange Online. Error occurred"
    }
}

# We are looking for the local profiles on this PC --- Start of my code
$path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
$profiles = (Get-ChildItem $path).PSChildName

# we need to map the HKU drive in-case it's not there by default
If(!(Test-Path HKU:)){
    Write-Output "Mapping HKU drive."
    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS
}
else{Write-Output "The HKU: drive is already mapped."}

#Used to track our progress
$progress = 0

# Check every key at HK local machine profile list
Foreach ($profileKey in $profiles)
{
    $value = "ProfileImagePath"
    $profileList = (Get-ItemProperty -Path $path\$profileKey -Name $value).$value

    # If the profilelist matches a local user we are updating the email signature there.
    if($profileList -like "C:\Users\*"){
        $localUser = $profileList.Split("\")[2]

        # starting to update outlook profile        
        Write-Output "Updating Outlook profile(s) for local user $localUser."
        
        # OutlookProfile paths
        $outlookProfilePath = "HKU:\$profileKey\Software\Microsoft\Office\16.0\Outlook\Profiles"

        #If profile doesn't exist, skip
        If(!(Test-Path $outlookProfilePath)){
            Write-Output "$outlookProfilePath does not exist, skipping"
            continue
        }

        # location of the signature generated
        $save_location = “C:\users\$localUser\Appdata\Roaming\Microsoft\Signatures”

        #If signature folder does not exist we need to create it
        If(!(Test-Path $save_location)){
            Write-Output "Creating Signatures Folder"
            New-Item -ItemType Directory -Force -Path $save_location
        }
        else{Write-Output "Signatures folder exist."}

        $outlookProfiles = (Get-ChildItem $OutLookProfilePath).PSChildName
        # Iterate through every available Outlook profile
        Foreach ($outlookProfile in $outlookProfiles)
        {
	        # Check every Mailbox.
            $mailboxPath = "$outlookProfilePath\$outlookProfile\9375CFF0413111d3B88A00104B2A6676"
	        $mailboxes = (Get-ChildItem $mailboxPath).PSChildName 
		    Foreach ($mailbox In $mailboxes ){
			    $OutlookMailboxPath = "$mailboxPath\$mailbox"    
			    $value = "Account Name"
			    $mailboxName = (Get-ItemProperty -Path $outlookMailboxPath -Name $value).$value

                If($mailboxName -match "mailbox we're trying to skip"){
                    Write-Output "$mailboxName does not exist, skipping"
                    continue
                }

			    # if the mailbox we are currently at has @domain, generate the signature for that mailbox and insert it.
			    If ($mailboxName -like "*@domain") {
				    $username = $mailboxName
                    Write-Output "Updating Signature for mailbox $username, located under profile $outlookProfile."

                    # Get user info and populate the signature 
                    $user = Get-User $username # this is where we put the user we are updating.
                    
                    #Signature information
                    $full_name = $($user.DisplayName)
                    # $account_name = $($user.UserPrincipalName) we don’t need account name for this script
                    $job_title = $($user.Title)
                    # $location = $($user.office) we don’t need location for now
                    $comp = “Company Name” # The company will always be Silver Gold Bull
                    $email = $($user.WindowsEmailAddress)
                    $phone = $($user.Phone)
                    # $logo = “C:/Path-to-photo” if they ever wanted to add logos you would have to update the HTML

                    # Generating the unique signature name
                    $sign = $mailboxName.split("@")[0]
                    $signatureFile = “signature-$sign.htm”
                    $signatureName = "signature-$sign"

                    $output_file = "$save_location\$signatureFile"

                    Write-Output "Creating HTML for $full_name."
                    $HTML = "<span style=`"font-family: calibri,sans-serif;`"><p style=`"margin: 0cm;`"><strong>$full_name<span style=`"color: #ffda1a;`"> | </span></strong>$job_title</p><p style=`"margin: 0cm;`">$comp</p><p style=`"margin: 0cm;`"><a href=`"mailto:$email`">$email</a></p><p style=`"margin: 0cm;`">$phone</p></span><br>"
                    $HTML | Out-File $output_file

                    #$OutlookProfilePath = "HKU:...\Software\Microsoft\Office\16.0\Common\MailSettings" This is where you go if you want to make sure they cannot change their signature settings.

                    # Insert the generated signature and WE ARE DONE!
				    Get-Item -Path $OutlookMailboxPath | New-Itemproperty -Name "New Signature" -value $signatureName -Propertytype string -Force 
				    Get-Item -Path $OutlookMailboxPath | New-Itemproperty -Name "Reply-Forward Signature" -value $signatureName -Propertytype string -Force
                    $progress += 1

			    }

		    }
        }

       
    }

}

if($progress -eq 0)
    {Write-Output "No Signatures were able to be updated"}
else
    {Write-Output "Updated $progress signatures.... Closing"}

Get-PSSession | Remove-PSSession
