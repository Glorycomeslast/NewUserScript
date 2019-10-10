#######################################################################
# 
# Must be run as administrator
# 
# Version 1.0
#
#######################################################################

If ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544"))
{

#######################################################################
#                Use this section to save credential                  #
#######################################################################
#
#$Username="#Enter Office 365 Login#"
#$Password= ConvertTo-SecureString "#Enter Password Here#" -AsPlainText -Force
#$psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)
#$Credential = $psCred
#
#######################################################################

$Credential = Get-Credential -ErrorAction Stop
$env:USERNAME

#######################################################################
#             Check AzureAD Module - Install If Missing               #
#######################################################################

$AzureAD = "AzureAD"

$Installedmodules = Get-InstalledModule

if ($Installedmodules.name -contains $AzureAD)
{

    "$AzureAD is installed "

}

else {

Install-Module AzureAD

"$AzureAD now installed"

}

#######################################################################
#              Check MSOnline Module - Install If Missing             #
#######################################################################

$MSOnline = "MSOnline"

$Installedmodules = Get-InstalledModule

if ($Installedmodules.name -contains $MSOnline)
{

    "$MSOnline is installed "

}

else {

Install-Module MSOnline

"$MSOnline now installed"

}

#######################################################################
#               Clear Screen and Connect MsolService                  #
#######################################################################

cls

Connect-MsolService -Credential $Credential -ErrorAction Stop

$Error.clear()

#######################################################################
#                           User Information                          #
#######################################################################

#Provide user's first name
$Firstname = Read-Host -Prompt "First Name"

#Provide user's last name
$Lastname = Read-Host -Prompt "Last Name"

#Get user's display name
$DisplayName = $Firstname + " " + $Lastname

#Provide user's username
$SAM = Read-Host -Prompt "Username"

#Provide user's department
$Department = Read-Host -Prompt "Department"

#Provide user's Azure region
$Location = Read-Host -Prompt "GB or US"

#Get user's User Principle Name
$UPN = $SAM + "{Enter Office365 Domain}"

#Default Address
$DefaultAddress = "#Enter Default Street Address#"
$DefaultCity = "#Enter Default City#"
$DefaultState = "#Enter Default County/State#"
$DefaultCountry = "#Enter Default Country#"
$DefaultPostalCode = "#Enter Default Postal Code#"

#Get users Street Address, if the input is left empty then it will automatically default to #Default Street Address#
$UserStreetAddress = (Read-Host -Prompt "Please input the users street address, will default to $DefaultAddress, please press enter if this is correct")

#Get users city, if the input is left empty then it will automatically default to #Default City#
$UserCity = (Read-Host -Prompt "Please input the users city, will default to $DefaultCity, please press enter if this is correct")

#Get users state, if the input is left empty then it will automatically default to #Default County/State#
$UserState = (Read-Host -Prompt "Please enter the county, if nothing is input it will default to' $DefaultState, please press enter if this is correct")

#Get users postal code, if the input is left empty then it will automatically default to #Default Postal Code#
$UserPostalCode = (Read-Host -Prompt "Please input the users Postal code, if left blank will default to $DefaultPostalCode, please press enter if this is correct")

#Get users country, if the input is left empty then it will automatically default to #Default Country#
$UserCountry = (Read-Host -Prompt "Please enter the country, if nothing is input this will default to $DefaultCountry, please press enter if this is correct")

#Ensure that user street address is not empty if it is, uses default address 
while ([string]::IsNullOrWhiteSpace($UserStreetAddress)) {$UserStreetAddress = $DefaultAddress}

#Ensure that user city is not empty, if it is uses default city
while ([string]::IsNullOrWhiteSpace($UserCity)) {$UserCity = $DefaultCity}

#Ensure that users state is not empty, if it is uses default state
while ([string]::IsNullOrWhiteSpace($UserState)) {$UserState = $DefaultState}

#Ensure that zip code is not empty, if not uses default value
while ([string]::IsNullOrWhiteSpace($UserPostalCode)) {$UserPostalCode = $DefaultPostalCode}

#Ensure country code is not empty, if it is use default country
while ([string]::IsNullOrWhiteSpace($UserCountry)) {$UserCountry = $DefaultCountry}

cls

Write-Host "Lets begin...."
Start-Sleep -Seconds 2

#######################################################################
#                          Creating The User                          #
#######################################################################

Write-Host "Creating user and mailbox"
New-MsolUser -DisplayName $DisplayName -FirstName $Firstname -LastName $Lastname -UserPrincipalName $UPN -UsageLocation $Location -Department $Department -ErrorAction Stop
Start-Sleep -Seconds 5

Write-Host "User and mailbox created!"

Write-Host "Adding address"
Set-MsolUser -UserPrincipalName $UPN -StreetAddress $UserStreetAddress -City $UserCity -State $UserState -PostalCode $UserPostalCode -Country $UserCountry -ErrorAction Stop
Start-Sleep -Seconds 5

Write-Host "Address added!"

#######################################################################
#                    Assigning Office 365 License                     #
#######################################################################

Write-Host "Applying license"
Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses #Specify Microsoft License# -ErrorAction Stop
Start-Sleep -Seconds 5

Write-Host "License applied!"

Start-Sleep -Seconds 5

Get-MsolUser -UserPrincipalName $UPN
}



else
{
    Write-Host
    
"#######################################################################
#                                                                     #
#                  MUST BE RUN AS ADMINISTRATOR                       #
#                                                                     #
#######################################################################"
}