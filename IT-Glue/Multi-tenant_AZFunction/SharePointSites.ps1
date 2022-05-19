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
$FlexAssetName = "Microsoft 365 SharePoint Sites"
$Description = "Documentation for all SharePoint Sites in Microsoft 365"

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
            icon        = 'globe'
            description = $description
        }
        relationships = @{
            "flexible-asset-fields" = @{
                data = @(
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order           = 1
                            name            = "Site Name"
                            kind            = "Text"
                            required        = $false
                            "show-in-list"  = $true
                            "use-for-title" = $true
                        }
                    },

                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 2
                            name           = "Site URL"
                            kind           = "Text"
                            required       = $false 
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 3
                            name           = "Last Activity Date"
                            kind           = "Date"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 4
                            name           = "Site File Count"
                            kind           = "Number"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 5
                            name           = "Site Storage Used-GB"
                            kind           = "Number"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 6
                            name           = "Storage Allocated-GB"
                            kind           = "Number"
                            required       = $false
                            "show-in-list" = $true
                        }
                    }
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 7
                            name           = "Microsoft Group Connected"
                            kind           = "Checkbox"
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
}while ($domains.count % 1000 -eq 0 -and $domains.count -ne 0)

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

$Sites = ""
    #####Get Site information if it is available####
try{
  $Sites = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getSharePointSiteUsageDetail(period='D30')" -Headers $headers -Method Get -ContentType "application/json") | convertfrom-csv  |select-object @{ Name = 'Site Name'; Expression = { $_.'Site URL'.split('/')[4] } },
  @{ Name = 'Site URL'; Expression = { $_.'Site URL' } },
  @{ Name = 'Site Last Activity Date'; Expression = { $_.'Last Activity Date' } },
  @{ Name = 'Site File Count'; Expression = { $_.'File Count' } },
  @{ Name = 'Site Storage Used (GB)'; Expression = { [math]::round($_.'Storage Used (Byte)' /1GB,0) } },
  @{ Name = 'Storage Allocated (GB)'; Expression = { [math]::round($_.'Storage Allocated (Byte)' /1GB,0) } },
  @{ Name = 'Microsoft Group connect'; Expression = { $_.'Root Web Template' } }
  }catch{('Error calling sites MS Graph')} 
  
  
  
  if($Sites){
  
    $SharePointObj = foreach ($site in $Sites) {
    
      $FlexAssetBody =
        @{
        type       = 'flexible-assets'
        attributes = @{
            traits = @{
        'site-name'                          = $site.'Site Name'
        "site-url"                           = $site.'Site URL'
        "last-activity-date"                 = $site.'Site Last Activity Date'
        'site-file-count'                    = $site.'Site File Count'
        "site-storage-used-gb"               = $site.'Site Storage Used (GB)'
        "storage-allocated-gb"               = $site.'Storage Allocated (GB)'
        "microsoft-group-connected"          = if($site.'Microsoft Group connect' -eq "Group"){$true}else{$false}
      }
    }
      
    }
  
      $ExistingFlexAsset = (Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $($filterID.id) -filter_organization_id $orgID).data | Where-Object { $_.attributes.traits.'site-url' -eq $site.'Site URL'}
          #If the Asset does not exist, we edit the body to be in the form of a new asset, if not, we just update.
          if (!$ExistingFlexAsset) {
              $FlexAssetBody.attributes.add('organization-id', $orgID)
              $FlexAssetBody.attributes.add('flexible-asset-type-id', $($filterID.ID))
              write-host "Creating Site in IT-Glue" -ForegroundColor Green
              New-ITGlueFlexibleAssets -data $FlexAssetBody
          }
          else {
              write-host "Updating Site in IT-Glue"  -ForegroundColor Yellow
              $ExistingFlexAsset = $ExistingFlexAsset | select-object -last 1
              Set-ITGlueFlexibleAssets -id $ExistingFlexAsset.id -data $FlexAssetBody
          }
      }
    }
  }}