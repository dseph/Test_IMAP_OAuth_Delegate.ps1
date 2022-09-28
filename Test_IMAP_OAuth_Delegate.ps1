#################################################################################################################################
# This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. # 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  #
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.               #
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object  #
# code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software   #
# product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the  #
# Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims   #
# or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.                 #
#################################################################################################################################

#----------------------------------------------------------------------              
#-     UPDATE VARIABLES TO REFLECT YOUR ENVIRONMENT                   -
#----------------------------------------------------------------------

# Provide Azure AD Application registration information for your app.
$AppID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"  
$TenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"  

#----------------------------------------------------------------------              
#-     DO NOT CHANGE ANY CODE BELOW THIS LINE                         -
#----------------------------------------------------------------------
#-                                                                    -
#-                           Author:  Dirk Buntinx                    -
#-                           Date:    12/9/2022                       -
#-                           Version: v1.0                            -
#-                                                                    -
#----------------------------------------------------------------------

Write-Host "--------------------"
Write-Host "- Script settings: -"
Write-Host "--------------------"
Write-Host " - AppID: $($AppID)"
Write-Host " - TenantId: $($TenantId)"
Write-Host "--------------------"
Write-Host "-  Start script    -"
Write-Host "--------------------"
Write-Host " - Retrieving OAuth token"

# Use the MSAL.PS PowerShell module to get the OAuth token
# install the MSAL.PS module from here: https://github.com/AzureAD/MSAL.PS
#Retrieve the MSAL Token
$Scopes = @("https://outlook.office365.com/IMAP.AccessAsUser.All")
$MSALToken = Get-MsalToken -ClientId $AppID -tenantID $TenantId -Scopes $Scopes -Interactive

Write-Host " - Importing PopImap Module"
# Use the PopImap module to connect to IMAP
# install the PopImap library version 0.1.3 from here: https://www.powershellgallery.com/packages/PopImap/0.1
# Import the module after installing
Import-Module PopImap 


$Server = "outlook.office365.com"
$Port = 993
Write-Host " - Connecting to IMAP server"
Write-Host " - Server: $($Server)" 
Write-Host " - Port: $($Port)" 
Write-Host " - Account: $($MSALToken.Account.UserName)" 
Write-Host ""

$imap = Get-ImapClient -Server $Server -Port $Port
$imap.Connect()
$imap.O365Authenticate($MSALToken.AccessToken, $MSALToken.Account.UserName)
$imap.ExecuteCommand("list `"INBOX/`" *")
$imap.Close()
Write-Host "--------------------"
Write-Host "-  End script      -"
Write-Host "--------------------"