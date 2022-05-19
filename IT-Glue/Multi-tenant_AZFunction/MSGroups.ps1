param($Timer)

Import-Module MsOnline -UseWindowsPowerShell
Import-Module ITGlueAPI -UseWindowsPowerShell

###MICROSOFT SECRETS#####

$ApplicationId = $ENV:ApplicationId
$ApplicationSecret = $ENV:ApplicationSecret
$tenantID = $ENV:tenantID
$refreshToken = $ENV:refreshToken
$secPas = $ApplicationSecret| ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $secPas)

########################## IT-Glue Information ############################
$ITGlueAPIKey = $ENV:ITGlueAPIKey
$APIEndpoint = "https://api.itglue.com"
$FlexAssetName = "Microsoft 365 Groups"
$Description = "Documentation for all groups in Microsoft 365"

#Settings IT-Glue logon information
Add-ITGlueBaseURI -base_uri $APIEndpoint
Add-ITGlueAPIKey $ITGlueAPIKEy

write-host "Checking if Flexible Asset exists in IT-Glue." -foregroundColor green
$FilterID = (Get-ITGlueFlexibleAssetTypes -filter_name $FlexAssetName).data
if (!$FilterID) { 
    write-host "Does not exist, creating new." -foregroundColor green
    $NewFlexAssetData = 
    @{
        type          = 'flexible-asset-types'
        attributes    = @{
            name        = $FlexAssetName
            icon        = 'users'
            description = $description
        }
        relationships = @{
            "flexible-asset-fields" = @{
                data = @(
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order           = 1
                            name            = "GroupName"
                            kind            = "Text"
                            required        = $true
                            "show-in-list"  = $true
                            "use-for-title" = $true
                        }
                    },

                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 2
                            name           = "Email"
                            kind           = "Text"
                            required       = $false 
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 3
                            name           = "Group Type"
                            kind           = "Text"
                            required       = $true
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 4
                            name           = "Membership Type"
                            kind           = "Text"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 5
                            name           = "Creation Date"
                            kind           = "Date"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 6
                            name           = "Source"
                            kind           = "Text"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 7
                            name           = "Security Enabled"
                            kind           = "Checkbox"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 8
                            name           = "Visibility"
                            kind           = "Text"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 9
                            name           = "HideFromGal"
                            kind           = "Checkbox"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 10
                            name           = "Teams Connected"
                            kind           = "Checkbox"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 11
                            name           = "Owners"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 12
                            name           = "Members"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                )
            }
        }
    }
    New-ITGlueFlexibleAssetTypes -Data $NewFlexAssetData
    $FilterID = (Get-ITGlueFlexibleAssetTypes -filter_name $FlexAssetName).data
}


#Grab all IT-Glue domains to match the domain name.
write-host "Getting IT-Glue domain list" -foregroundColor green
$i = 1
$AllITGlueDomains = do {
    $domains = (Get-ITGlueDomains -page_size 1000 -page_number $i).data.attributes
    $i++
    $domains
    Write-Host "Retrieved $($domains.count) Domains" -ForegroundColor Yellow
}while ($Contacts.count % 1000 -eq 0 -and $Contacts.count -ne 0)

$DomainList = foreach ($domain in $AllITGlueDomains) {
    $ITGDomain = $domain.'name'
    [PSCustomObject]@{
        Domain   = $ITGDomain
        OrgID    = $domain.'organization-id'
        Combined = "$($ITGDomain)$($domain.'organization-id')"
    }
}


###Connect to your Own Partner Center to get a list of customers/tenantIDs #########
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID


Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken

$customers = Get-MsolPartnerContract -All
 
Write-Host "Found $($customers.Count) customers in Partner Center." -ForegroundColor DarkGreen

foreach ($customer in $customers) {
  Write-Host "Found $($customer.Name) in Partner Center" -ForegroundColor Green

  $CustomerDomains = Get-MsolDomain -TenantId $customer.TenantID
  $orgid = foreach ($customerDomain in $customerdomains) {
      ($domainList | Where-Object { $_.domain -eq $customerDomain.name }).'OrgID'
  }

  $orgID = $orgid | Select-Object -Unique
  if(!$orgID){
     Write-Host "Customer does not exist in IT-Glue" -ForegroundColor Red
  }
  if($orgID){

    ###Get Access Token########
    $CustomerToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -Tenant $customer.TenantID
    $headers = @{ "Authorization" = "Bearer $($CustomerToken.AccessToken)" }
    
    #####Get Group information if it is available####
    try{
      $Groups = (Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/groups' -Headers $headers -Method Get -ContentType "application/json").value | Select-Object id, displayName, mail, groupTypes,createdDateTime,onPremisesSyncEnabled,securityEnabled,visibility,resourceBehaviorOptions,resourceProvisioningOptions
      }catch{('Error calling devices MS Graph')} 
  
  
    foreach ($group in $Groups) {
  
        ##Get Owners
        try{
          $Owners = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/owners" -Headers $headers -Method Get -ContentType "application/json").value | Select-Object displayName, mail
          }catch{('Error calling devices MS Graph')} 
          
        ##Get Members
        try{
          $Members = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/members" -Headers $headers -Method Get -ContentType "application/json").value | Select-Object displayName, mail
          }catch{('Error calling devices MS Graph')}  
  
        ##Determine Group Type
        if($group.groupTypes -contains "Unified"){
          $GroupType = "Microsoft 365"
        } elseif( (!$group.groupTypes) -and $group.securityEnabled -eq $false){
          $GroupType = "Distribution"
        } else{
          $GroupType = "Security"
        }
        ##Determine Dynamic Or Assigned Membership
        if($group.groupTypes -contains "DynamicMembership"){
          $membership = "Dynamic"
        } else{
          $membership = "Assigned"
        }
        ##Determine Hide from GAL
        if($group.resourceBehaviorOptions -contains "HideGroupInOutlook"){
          $hidefromGal = $true
        } else{
          $hidefromGal = $false
        }
  
        ##Determine if Teams Channel Assoicated
        if($group.resourceProvisioningOptions -contains "Team"){
          $teamConnected = $true
        } else{
          $teamConnected = $false
        }
  
  
        $FlexAssetBody =
        @{
        type       = 'flexible-assets'
        attributes = @{
            traits = @{
          'groupname'                          = $group.displayName
          "email"                              = if($group.mail){$group.mail}else{"N/A"}
          "group-type"                         = $GroupType
          'membership-type'                    = $membership
          "creation-date"                      = $group.createdDateTime
          "source"                             = if($group.onPremisesSyncEnabled){"On-premise"}else{"Cloud"}
          "security-enabled"                   = $group.securityEnabled
          "visibility"                         = $group.visibility
          "hidefromgal"                        = $hidefromGal
          "teams-connected"                    = $teamConnected
          "owners"                             = ($Owners | Select-Object displayName, mail | convertto-html -Fragment  | out-string)
          "members"                            = ($Members | Select-Object displayName, mail | convertto-html -Fragment  | out-string)
            }
      }
    }

  
      $ExistingFlexAsset = (Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $($filterID.id) -filter_organization_id $orgID).data | Where-Object { $_.attributes.traits.'groupname' -eq $group.displayName}
          #If the Asset does not exist, we edit the body to be in the form of a new asset, if not, we just update.
          if (!$ExistingFlexAsset) {
              $FlexAssetBody.attributes.add('organization-id', $orgID)
              $FlexAssetBody.attributes.add('flexible-asset-type-id', $($filterID.ID))
              write-host "Creating Group: $($group.displayName) in IT-Glue" -ForegroundColor Green
              New-ITGlueFlexibleAssets -data $FlexAssetBody
          }
          else {
              write-host "Updating Group: $($group.displayName) in IT-Glue"  -ForegroundColor Yellow
              $ExistingFlexAsset = $ExistingFlexAsset | select-object -last 1
              Set-ITGlueFlexibleAssets -id $ExistingFlexAsset.id -data $FlexAssetBody
          }
    }
    }}