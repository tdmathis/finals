############################################################################### 
# 
# Function ConnectTo-ExMSOnline 
# 
# PURPOSE 
#    Connects to MSOnline & Exchange Online Remote PowerShell using admin credentials 
# 
# INPUT 
#    Admin username and password. 
# 
# RETURN 
#    None. 
# 
############################################################################### 
function ConnectTo-ExMSOnline {    
    
    Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string]$O365AdminUsername, 
    [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)] [string]$O365AdminPassword
    )
    
    #Remove all existing Powershell sessions 
    Get-PSSession | Remove-PSSession 
    
    #Encrypt password for transmission to Office365 
    $SecureO365Password = ConvertTo-SecureString -AsPlainText $O365AdminPassword -Force     
     
    #Build credentials object
    #$LiveCred = Get-Credential 
    $LiveCred  = New-Object System.Management.Automation.PSCredential $O365AdminUsername, $SecureO365Password 
     
    #Create remote Powershell session 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $LiveCred -Authentication Basic –AllowRedirection  
 
    #Import the session 
    Import-PSSession $Session #-AllowClobber | Out-Null: -AllowClobber is useful when working from a server that is hosting an exechange server, | Out-Null will delete the redirection messages
    
    #Connect to MSOnline to view & assign licenses (relies on Import-Module MSOnline from main)
    Connect-MSOLservice -Credential $LiveCred 
} 

############################################################################### 
# 
# Function New-MSOnlineO365User 
# 
# PURPOSE 
#    Tests the requested username against length and complexity requirements, creates the account, assigns the license 
# 
# INPUT 
#    Email address, given name, surname(optional), domain(optional), license type(optional)
# 
# RETURN 
#    None
# 
############################################################################### 
function New-MSOnlineO365User {
    
    #change to read-host?
    #Accept input parameters 
    Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string]$O365EmailAddress,
    [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)] [string]$O365gn,
    [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] [string]$O365sn, 
    [Parameter(Position=3, Mandatory $false, ValueFromPipeline=$true)] [string]$O365domain, 
    [Parameter(Position=4, Mandatory=$false, ValueFromPipeline=$true)] [string]$LicenseType
    ) 
    
    Import-Module MSOnline
    
    $O365EmailAddress = $O365EmailAddress.toLower()
    $O365Username = $O365EmailAddress.Split("@")[0].toLower()
    
    #If no $O365sn is passed, create one from the username #$sn = $sn.Replace($sn[0],$sn[0].ToString().ToUpper())
    if ($O365sn) {} 
    else {$O365sn = $O365Username.substring(1,2).toUpper()+$O365Username.substring(2).tolower()} 
    
    #If no $O365domain is passed, create one from the email address
    if ($O365domain) {}
    else { $O365domain = $O365EmailAddress.Split("@")[1]}
    
    #If no $LicenseType is passed, assume EXCHANGESTANDARD #O365:EXCHANGESTANDARD  Federated:EXCHANGEENTERPRISE
    if ($LicenseType) { $LicenseType = "EXCHANGEENTERPRISE"}
    else { $LicenseType = "EXCHANGESTANDARD"}
    
    #checks if the username matches requirements
    $UsernameRegex = "^[A-Za0-9-_.]+$" #Username may contantain alpha-numeric characters, dot, hyphen and/or underscore 
    do {
        
        #This section needs testing ##############################################################################
        Try {
        
            #Checks to see if a mailbox exsists on the server with the chosen username
            $MailboxExsists = Get-Mailbox -Filter {sAMAccountName -eq $O365Username} 
             } 
        
        #If error, the username has not been used (creates an error because EMC can't find the mailbox)
        Catch [system.exception]{ $Error }
        
        Finally { 
            If $Error { # continue silently #$Error | fl * -f 
                }
            
            Else { 
                
                #Display mailboxes with first two characters of the selected username, expanded with first and last name
                while ($MailboxExsists -ne $Null) { 
                
                    Get-Mailbox $O365Username.Substring(0,2)* | select -expand name 
                    $O365Username = Read-host "That username is unavailble, please select another:"
                    }
                }  
        ################################################################################################################
        
        #Username must be at least three characthers long
        if($O365Username.length -lt 3){  
            
            $O365Username = Read-host "That username isn't long enough, please select another."
            }
        
        #Username must match the regular expression defined in Variables
        elseif ($O365Username -notmatch $UsernameRegex){
        
            $O365Username = Read-host "That user name does not conform, please select another."
            }
            
        #Silently passes validated username
        else {}
        
        #Loops until username meets requirements 
        } while ($O365Username.length -lt 3 -OR $O365Username -notmatch $UsernameRegex -OR $MailboxExsists -ne $Null) # test this last part
   
   
   New-MsolUser -UserPrincipalName $O365Username -DisplayName "$O365gn $O365sn" -UsageLocation US -LicenseAssignment "$O365domain:$LicenceType"
    
    }

############################################################################### 
# 
# Function Add-MailboxAlias 
# 
# PURPOSE 
#    Adds proxy email addresses to Exchange mailboxes 
# 
# INPUT 
#    Username, array of domains, manual alias
# 
# RETURN 
#    None
# 
###############################################################################    
function Add-MailboxAlias{
    
    #Accept input parameters 
    Param( 
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [string]$Username, 
    [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)] [array]$Domains,
    [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)] [string]$altAlias
    ) 
    
    #Update the mail policy to allow aliasing #needs testing
    $Username.EmailAddressPolicyEnabled = $false
    
    #Add the alias's
    foreach ($alias in $Domains) {
        $Username.EmailAddresses += $Username + "@" + $alias + ".com, " #this needs to be tested/modded
        }
    
    #add aliases with different usernames #this needs to be tested
    if $altAlias {$Username.EmailAddresses += $altAlias }
    
    #display results
    get-mailbox $Username | select -expand Emailaddresses alias
    
    }
