#Export all Customer's Office 365 Licences to individual .csv files
#John Lanigan 2022

#Establish a PowerShell session with Office 365. You'll be prompted for your Delegated Admin credentials

Connect-MsolService

$LicenseLookup = @{
    'SPZA_IW'                                 = 'App Connect Iw'
    'AAD_BASIC'                               = 'Azure Active Directory Basic'
    'AAD_PREMIUM'                             = 'Azure Active Directory Premium P1'
    'AAD_PREMIUM_P2'                          = 'Azure Active Directory Premium P2'
    'RIGHTSMANAGEMENT'                        = 'Azure Information Protection Plan 1'
    'MCOCAP'                                  = 'Common Area Phone'
    'MCOPSTNC'                                = 'Communications Credits'
    'DYN365_ENTERPRISE_PLAN1'                 = 'Dynamics 365 Customer Engagement Plan Enterprise Edition'
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE'      = 'Dynamics 365 For Customer Service Enterprise Edition'
    'DYN365_FINANCIALS_BUSINESS_SKU'          = 'Dynamics 365 For Financials Business Edition'
    'DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE' = 'Dynamics 365 For Sales And Customer Service Enterprise Edition'
    'DYN365_ENTERPRISE_SALES'                 = 'Dynamics 365 For Sales Enterprise Edition'
    'DYN365_ENTERPRISE_TEAM_MEMBERS'          = 'Dynamics 365 For Team Members Enterprise Edition'
    'DYN365_TEAM_MEMBERS'                     = 'Dynamics 365 Team Members'
    'Dynamics_365_for_Operations'             = 'Dynamics 365 Unf Ops Plan Ent Edition'
    'EMS'                                     = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                              = 'Enterprise Mobility + Security E5'
    'EXCHANGESTANDARD'                        = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                      = 'Exchange Online (Plan 2)'
    'EXCHANGEARCHIVE_ADDON'                   = 'Exchange Online Archiving For Exchange Online'
    'EXCHANGEARCHIVE'                         = 'Exchange Online Archiving For Exchange Server'
    'EXCHANGEESSENTIALS'                      = 'Exchange Online Essentials'
    'EXCHANGE_S_ESSENTIALS'                   = 'Exchange Online Essentials'
    'EXCHANGEDESKLESS'                        = 'Exchange Online Kiosk'
    'EXCHANGETELCO'                           = 'Exchange Online Pop'
    'INTUNE_A'                                = 'Intune'
    'M365EDU_A1'                              = 'Microsoft 365 A1'
    'M365EDU_A3_FACULTY'                      = 'Microsoft 365 A3 For Faculty'
    'M365EDU_A3_STUDENT'                      = 'Microsoft 365 A3 For Students'
    'M365EDU_A5_FACULTY'                      = 'Microsoft 365 A5 For Faculty'
    'M365EDU_A5_STUDENT'                      = 'Microsoft 365 A5 For Students'
    'O365_BUSINESS'                           = 'Microsoft 365 Apps For Business'
    'SMB_BUSINESS'                            = 'Microsoft 365 Apps For Business'
    'OFFICESUBSCRIPTION'                      = 'Microsoft 365 Apps For Enterprise'
    'MCOMEETADV'                              = 'Microsoft 365 Audio Conferencing'
    'MCOMEETADV_GOC'                          = 'Microsoft 365 Audio Conferencing For Gcc'
    'O365_BUSINESS_ESSENTIALS'                = 'Microsoft 365 Business Basic'
    'SMB_BUSINESS_ESSENTIALS'                 = 'Microsoft 365 Business Basic'
    'SPB'                                     = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'                   = 'Microsoft 365 Business Standard'
    'SMB_BUSINESS_PREMIUM'                    = 'Microsoft 365 Business Standard'
    'MCOPSTN_5'                               = 'Microsoft 365 Domestic Calling Plan (120 Minutes)'
    'SPE_E3'                                  = 'Microsoft 365 E3'
    'SPE_E3_USGOV_DOD'                        = 'Microsoft 365 E3_Usgov_Dod'
    'SPE_E3_USGOV_GCCHIGH'                    = 'Microsoft 365 E3_Usgov_Gcchigh'
    'SPE_E5'                                  = 'Microsoft 365 E5'
    'INFORMATION_PROTECTION_COMPLIANCE'       = 'Microsoft 365 E5 Compliance'
    'IDENTITY_THREAT_PROTECTION'              = 'Microsoft 365 E5 Security'
    'IDENTITY_THREAT_PROTECTION_FOR_EMS_E5'   = 'Microsoft 365 E5 Security For Ems E5'
    'M365_F1'                                 = 'Microsoft 365 F1'
    'SPE_F1'                                  = 'Microsoft 365 F3'
    'M365_G3_GOV'                             = 'Microsoft 365 Gcc G3'
    'MCOEV'                                   = 'Microsoft 365 Phone System'
    'PHONESYSTEM_VIRTUALUSER'                 = 'Microsoft 365 Phone System - Virtual User'
    'MCOEV_DOD'                               = 'Microsoft 365 Phone System For Dod'
    'MCOEV_FACULTY'                           = 'Microsoft 365 Phone System For Faculty'
    'MCOEV_GOV'                               = 'Microsoft 365 Phone System For Gcc'
    'MCOEV_GCCHIGH'                           = 'Microsoft 365 Phone System For Gcchigh'
    'MCOEVSMB_1'                              = 'Microsoft 365 Phone System For Small And Medium Business'
    'MCOEV_STUDENT'                           = 'Microsoft 365 Phone System For Students'
    'MCOEV_TELSTRA'                           = 'Microsoft 365 Phone System For Telstra'
    'MCOEV_USGOV_DOD'                         = 'Microsoft 365 Phone System_Usgov_Dod'
    'MCOEV_USGOV_GCCHIGH'                     = 'Microsoft 365 Phone System_Usgov_Gcchigh'
    'WIN_DEF_ATP'                             = 'Microsoft Defender Advanced Threat Protection'
    'CRMSTANDARD'                             = 'Microsoft Dynamics Crm Online'
    'CRMPLAN2'                                = 'Microsoft Dynamics Crm Online Basic'
    'FLOW_FREE'                               = 'Microsoft Flow Free'
    'INTUNE_A_D_GOV'                          = 'Microsoft Intune Device For Government'
    'POWERAPPS_VIRAL'                         = 'Microsoft Power Apps Plan 2 Trial'
    'TEAMS_FREE'                              = 'Microsoft Team (Free)'
    'TEAMS_EXPLORATORY'                       = 'Microsoft Teams Exploratory'
    'IT_ACADEMY_AD'                           = 'Ms Imagine Academy'
    'ENTERPRISEPREMIUM_FACULTY'               = 'Office 365 A5 For Faculty'
    'ENTERPRISEPREMIUM_STUDENT'               = 'Office 365 A5 For Students'
    'EQUIVIO_ANALYTICS'                       = 'Office 365 Advanced Compliance'
    'ATP_ENTERPRISE'                          = 'Microsoft Defender for Office 365 (Plan 1)'
    'STANDARDPACK'                            = 'Office 365 E1'
    'STANDARDWOFFPACK'                        = 'Office 365 E2'
    'ENTERPRISEPACK'                          = 'Office 365 E3'
    'DEVELOPERPACK'                           = 'Office 365 E3 Developer'
    'ENTERPRISEPACK_USGOV_DOD'                = 'Office 365 E3_Usgov_Dod'
    'ENTERPRISEPACK_USGOV_GCCHIGH'            = 'Office 365 E3_Usgov_Gcchigh'
    'ENTERPRISEWITHSCAL'                      = 'Office 365 E4'
    'ENTERPRISEPREMIUM'                       = 'Office 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF'            = 'Office 365 E5 Without Audio Conferencing'
    'DESKLESSPACK'                            = 'Office 365 F3'
    'ENTERPRISEPACK_GOV'                      = 'Office 365 Gcc G3'
    'MIDSIZEPACK'                             = 'Office 365 Midsize Business'
    'LITEPACK'                                = 'Office 365 Small Business'
    'LITEPACK_P2'                             = 'Office 365 Small Business Premium'
    'WACONEDRIVESTANDARD'                     = 'Onedrive For Business (Plan 1)'
    'WACONEDRIVEENTERPRISE'                   = 'Onedrive For Business (Plan 2)'
    'POWER_BI_STANDARD'                       = 'Power Bi (Free)'
    'POWER_BI_ADDON'                          = 'Power Bi For Office 365 Add-On'
    'POWER_BI_PRO'                            = 'Power Bi Pro'
    'PROJECTCLIENT'                           = 'Project For Office 365'
    'PROJECTESSENTIALS'                       = 'Project Online Essentials'
    'PROJECTPREMIUM'                          = 'Project Online Premium'
    'PROJECTONLINE_PLAN_1'                    = 'Project Online Premium Without Project Client'
    'PROJECTPROFESSIONAL'                     = 'Microsoft Project Plan 3'
    'PROJECTONLINE_PLAN_2'                    = 'Project Online With Project For Office 365'
    'SHAREPOINTSTANDARD'                      = 'Sharepoint Online (Plan 1)'
    'SHAREPOINTENTERPRISE'                    = 'Sharepoint Online (Plan 2)'
    'MCOIMP'                                  = 'Skype For Business Online (Plan 1)'
    'MCOSTANDARD'                             = 'Skype For Business Online (Plan 2)'
    'MCOPSTN2'                                = 'Skype For Business Pstn Domestic And International Calling'
    'MCOPSTN1'                                = 'Skype For Business Pstn Domestic Calling'
    'MCOPSTN5'                                = 'Skype For Business Pstn Domestic Calling (120 Minutes)'
    'MCOPSTNEAU2'                             = 'Telstra Calling For O365'
    'TOPIC_EXPERIENCES'                       = 'Topic Experiences'
    'VISIOONLINE_PLAN1'                       = 'Visio Online Plan 1'
    'VISIOCLIENT'                             = 'Visio Online Plan 2'
    'VISIOCLIENT_GOV'                         = 'Visio Plan 2 For Gov'
    'WIN10_PRO_ENT_SUB'                       = 'Windows 10 Enterprise E3'
    'WIN10_VDA_E3'                            = 'Windows 10 Enterprise E3'
    'WIN10_VDA_E5'                            = 'Windows 10 Enterprise E5'
    'WINDOWS_STORE'                           = 'Windows Store For Business'
    'RMSBASIC'                                = 'Azure Information Protection Basic'
    'UNIVERSAL_PRINT_M365'                    = 'Universal Print'
    'RIGHTSMANAGEMENT_ADHOC'                  = 'Rights Management Service Basic Content Protection'
    'SKU_Dynamics_365_for_HCM_Trial'          = 'Dynamics 365 for Talent'
    'PROJECT_P1'                              = 'Project Plan 1'
    'PROJECT_PLAN1_DEPT'                      = 'Project Plan  1 (Self Service)'
    'SHAREPOINTSTORAGE'                       = 'Microsoft Office 365 Extra File Storage'
    'NONPROFIT_PORTAL'                        = 'Non Profit Portal'
    'MDE_SMB'                                 = 'Microsoft Defender for Endpoint (Business Premium)'
}


$Customers = Get-MsolPartnerContract -All
Write-Host "Found $($Customers.Count) customers for $((Get-MsolCompanyInformation).displayname)." -ForegroundColor DarkGreen

  
foreach ($Customer in $Customers) 
{
    Write-Host "Retrieving license info for $($Customer.name)" -ForegroundColor Green
    $CSVpath = "C:\Temp\O365\$($Customer.name).csv"
    $AccountSKUs = Get-MsolAccountSku -TenantId $Customer.TenantId
    
    foreach ($SKU in $AccountSKUs) 
    {
        $license = $SKU.AccountSkuId
        $index = $license.indexof(":")
        $remove=$license.SubString(0,$index+1)
        $output = $license.Replace($remove,"").Replace(" ","")
        
        foreach ($pair in $licenselookup.GetEnumerator())
        {
            $output= $output -replace $pair.Name, $pair.Value
        }
        
        $i=1
        #get the users in this account that have this SKU
        $licensedUsers = Get-MsolUser -TenantId $customer.TenantId -All | Where-Object {($_.licenses).AccountSkuId -match $SKU.AccountSkuId}
        foreach ($User in $licensedUsers)
        {
                             
            $licenceinfo = [pscustomobject][ordered]@{
                License    = $output
                Purchased  = $SKU.ActiveUnits
                Assigned   = $SKU.ConsumedUnits
                Count      = $i
                User       = $user.DisplayName
             }
            $i=$i+1
            $licenceinfo | Export-CSV -Path $CSVpath -Append -NoTypeInformation 
            
        }
    $licenceinfo = [pscustomobject][ordered]@{
        License    = "-------"
        Purchased  = "-------"
        Assigned   = "-------"
        Count      = "-------"
        User       = "-------"
        }
     $licensedSharedMailboxProperties | Export-CSV -Path $CSVpath -Append -NoTypeInformation 
    }
}