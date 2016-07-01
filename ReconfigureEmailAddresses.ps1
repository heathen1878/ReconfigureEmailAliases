# Import the necessary functions
. .\Functions.ps1

# Start of script.
WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Script started...$(Get-Date)"

# Declare variables
[bool]$Validator = $false
[string]$sQuery = ""
[array]$bValid = @()
[System.Collections.ArrayList]$newEmailAddresses = @()
[string]$RemoveFromArray = ""

## Verification of the domain should be tested by connecting to Microsoft Online
While (!$validator) {
	
	# Get credentials for Office 365.
	$Creds = Get-Credential -Message "Please enter your Office 365 credentials"
	$validator = ConnectToMicrosoftOnline -credentials $creds

}

# Reset the validator
[bool]$Validator = $false

# Get the new domain name from the console
While (!$validator) {

	$newPrimaryDomain = Read-Host "What domain will be the new primary SMTP domain?"	
	If (Get-MsolDomain -DomainName $newPrimaryDomain -ErrorAction SilentlyContinue) {
		$Validator = $True
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$newPrimaryDomain validated"
	}
	else {
		Write-Output "$newPrimaryDomain not valid"
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$newPrimaryDomain not valid"
	}
}

# Reset the validator
[bool]$Validator = $false

# Get the new domain name from the console
While (!$validator) {

	$currentPrimaryDomain = Read-Host "What domain is being replaced?"
	If (Get-MsolDomain -DomainName $currentPrimaryDomain -ErrorAction SilentlyContinue) {
		$Validator = $True
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$currentPrimaryDomain validated"
	}
	else {
		Write-Output "$currentPrimaryDomain not valid"
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$currentPrimaryDomain not valid"
	}
}

# Reset the validator
[bool]$Validator = $false

# Set validator
$bValid = "Place holder"
$bValid += $false

# Call connect to Exchange Online
While (!$bValid[1]) {
	
	$bValid = ConnectToExchange -Credentials $creds -Location "R"
	If (!$bValid[1]){
		$Creds = Get-Credential -Message "Please enter your Office 365 credentials"	
	}	
	Else {
		Import-PSSession $bValid[0]
	}
}

# Create a query for the Get-Mailbox cmdlet.
$sQuery = -Join ("*@",$currentPrimaryDomain)

# Get a collection of mailboxes which have the current domain has a primary SMTP address. This will work with regular mailboxes, resource mail and shared mailboxes.
Get-mailbox -ResultSize Unlimited | Where-Object {$_.PrimarySmtpAddress -like $sQuery} | ForEach-Object {
		
	Try 
	{	
		# required for PowerShell 4.0 remoting.
		$Global:ErrorActionPreference="Stop"
		
		$UserAlias = $_.Alias
		
		# Get the current user in the pipeline into a variable 
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Processing user: $_...$(Get-Date)"

		# convert all instances of SMTP to smtp - believe it or not that is the difference
		# between primary and alias email addresses. 
		$newEmailAddresses = $_.EmailAddresses.Replace("SMTP","smtp")

        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Email address collection: $newEmailAddresses"
		
		# Get the existing primary SMTP address and use it to create the new primary SMTP address
		# e.g. user@domain.com to user@domain.uk...
		$newPrimarySMTPAddress = $_.PrimarySmtpAddress.Replace($currentPrimaryDomain,$newPrimaryDomain)
		
        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New Email address: $newPrimarySMTPAddress"

        # Remove the instance of the domain being added
        
        # Get the primary domain which exists in the array
        ### Need to check if it exists, if not then the add primary domain will suffice. ####
        $newEmailAddresses | ForEach-Object {
        
            If ($_ -match $newPrimaryDomain){
                
                $RemoveFromArray = $_
        
            }

        }

                Write-Output $RemoveFromArray

        # Check whether anything needs removing from the array
        If (-not $RemoveFromArray -eq ''){

            WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Removing: $RemoveFromArray"
    
            [void]$newEmailAddresses.Remove($RemoveFromArray)

        }

		# Make the new SMTP address primary by prefixing it with a uppercase SMTP.
		$newPrimarySMTPAddress = -Join ("SMTP:",$newPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New primary SMTP address will be $newPrimarySMTPAddress"
		[void]$newEmailAddresses.Add($newPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Addresses to be configured: $newEmailAddresses"
		
		# Update the email addresses associated with the mailbox
		Set-Mailbox -Identity $userAlias -EmailAddresses $newEmailAddresses
		
		# Check our work.
		Write-Output "Checking the configuration" -foregroundcolor Yellow
		Get-mailbox -Identity $UserAlias | select-Object PrimarySmtpAddress
		
	}		
	Catch 
	{
		
		If ($error[0].exception.message -match "is already present in the collection"){
			Write-Output "Error caught. Check the log file for more information." -foregroundcolor Yellow
			WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$Error[0].Exception.Message"
		}
        Else {
            WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$Error[0].Exception.Message"
        }

		
	}
	Finally
	{
		clear-variable UserAlias
		clear-variable newPrimarySMTPAddress
		clear-variable newEmailAddresses
        clear-Variable RemoveFromArray
		$Global:ErrorActionPreference="Continue"
		$Error.Clear()
	}
}

# Get a collection of distribution groups which have the current domain has a primary SMTP address.
Get-DistributionGroup -ResultSize Unlimited | Where-Object {$_.PrimarySmtpAddress -like $sQuery} | ForEach-Object {
		
	Try 
	{	
		# required for PowerShell 4.0 remoting.
		$Global:ErrorActionPreference="Stop"
		
		$UserAlias = $_.Alias
		
		# Get the current user in the pipeline into a variable
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Processing distribution group: $_...$(Get-Date)"
        
        # convert all instances of SMTP to smtp - believe it or not that is the difference
		# between primary and alias email addresses. 
		$newEmailAddresses = $_.EmailAddresses.Replace("SMTP","smtp")

        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Email address collection: $newEmailAddresses"
		
		# Get the existing primary SMTP address and use it to create the new primary SMTP address
		# e.g. user@domain.com to user@domain.uk...
		$newPrimarySMTPAddress = $_.PrimarySmtpAddress.Replace($currentPrimaryDomain,$newPrimaryDomain)
		
        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New Email address: $newPrimarySMTPAddress"

        # Remove the instance of the domain being added
        
        # Get the primary domain which exists in the array
        ### Need to check if it exists, if not then the add primary domain will suffice. ####
        $newEmailAddresses | ForEach-Object {
        
            If ($_ -match $newPrimaryDomain){
                
                $RemoveFromArray = $_
        
            }

        }

                Write-Output $RemoveFromArray

        # Check whether anything needs removing from the array
        If (-not $RemoveFromArray -eq ''){

            WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Removing: $RemoveFromArray"
    
            # Don't return update to screen
            [void]$newEmailAddresses.Remove($RemoveFromArray)

        }

		# Make the new SMTP address primary by prefixing it with a uppercase SMTP.
		$newPrimarySMTPAddress = -Join ("SMTP:",$newPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New primary SMTP address will be $newPrimarySMTPAddress"
		
        # Don't return update to screen
        [void]$newEmailAddresses.Add($newPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Addresses to be configured: $newEmailAddresses"
				
		# Update the email addresses associated with the mailbox
		Set-DistributionGroup -Identity $userAlias -EmailAddresses $newEmailAddresses
		
		# Check our work.
		Write-Output "Checking the configuration" -foregroundcolor Yellow
		Get-DistributionGroup -Identity $UserAlias | select-Object PrimarySmtpAddress
		
	}		
	Catch 
	{
		# Any error generated in the Try block are written to a log file in this Catch block.
		If ($error[0].exception.message -match "is already present in the collection"){
			Write-Output "Error caught. Check the log file for more information." -foregroundcolor Yellow
			WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$Error[0].Exception.Message"
		}
		
	}
	Finally
	{
		clear-variable UserAlias
		clear-variable newPrimarySMTPAddress
		clear-variable newEmailAddresses
        Clear-Variable RemoveFromArray
		$Global:ErrorActionPreference="Continue"
		$Error.Clear()
	}
}

# Clean up the remote PowerShell session created earlier.
Get-PSSession | Remove-PSSession