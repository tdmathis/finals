#
# Itemizes Active Directory users, hosts and servers by iterating 
# through groups by location, then exports numerations to .csv
#
# Takes one optional parameter, $locationCode, which will enumerate
# only that location's group
#
 

#allows you to limit the search to one location
param([string] $locationCode)
if($locationCode) {$franchiseCode = "ou=" + $locationCode + ",ou=Franchise,dc=americas,dc=franchise,dc=local"}
else { $hotels = "ou=Franchise,dc=americas,dc=franchise,dc=local" }

#a blank array for output and an empty hash table to allocate space for value
$template = New-Object psobject -Property @{
#$tempalte = [pscustomobject]@{
  OU=$null
  OpsUsers=$null
  FDUsers=$null
  Servers=$null
  Computers=$null
  }
$objTemp = $template | Select-Object *
$objResult = @()

#returns the children of $franchiseCode OUs
Get-ADOrganizationalUnit -filter * -SearchBase $locationCode -SearchScope Subtree | 
 foreach {

    $dn = $_.distinguishedname
    #filters a few test OUs
    while(($dn -like "*TEST1*") -or($dn -like "*TEST2*")-or($dn -like "*TEST3*")){ return; }

    #counts items depending on container
    switch ($_.Name) {
      "Users" {
             $objTemp = $objTemp | Select-Object *

             #selects the OU
             $ou = ($dn).Substring(12,5)
             $objTemp.OU = $ou

             #uses the OU to find enabled users in their groups
             $opsGroup = $ou +" OPS Users"
             $fdGroup = $ou +" FD Users"
             $opsEnabled = (Get-ADUser -LdapFilter "(&(!useraccountcontrol:1.2.840.113556.1.4.803:=2)(memberof=$(Get-ADGroup $opsGroup)))").count | out-string
             $fdEnabled = (Get-ADUser -LdapFilter "(&(!useraccountcontrol:1.2.840.113556.1.4.803:=2)(memberof=$(Get-ADGroup $fdGroup)))").count | out-string

             #this handles an issue where the LDAP filter was not returning 1
             if (($opsEnabled - '0') -lt 0){ $objTemp.OpsUsers = 1;}    
             else { $objTemp.OpsUsers = $opsEnabled -as [int]}      
   
             if (($fdEnabled - '0') -lt 0){$objTemp.FDUsers = 1;}
             else {$objTemp.FDUsers = $fdEnabled -as [int]}
            
             #adds the location to total results and starts a new temp object
             $objResult += $objTemp
             $objTemp = $template | Select-Object * 
            }
       "Computers" {
             $objTemp = $objTemp | Select-Object *

             #counts computers
             $c=Get-ADcomputer -filter * -searchbase $dn
             $objTemp.Computers = ($c | measure-object).count
            }
        "Servers" {
            $objTemp = $objTemp | Select-Object *

            #counts servers
            $s=Get-ADcomputer -filter * -searchbase $dn
            $objTemp.Servers=($s | measure-object).count
           }
        default { <#take no action#> }
      }
  }
  
  #if the temp object has not been populated, do not append to results
  if ($objTemp.OU -eq $null) {}
  else{ $objResult += $objTemp }

  #write results to terminal and export to .csv
  $objResult
  $objResult | Export-Csv c:\Users\$env:USERNAME\Documents\"IHG-FranchiseHotels"$locationCode.csv