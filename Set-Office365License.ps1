Function Set-Office365License{
    [CmdletBinding()]
    param(
    [parameter(mandatory=$true,Position=0)][string]$UserPrincipalName,
    [parameter(mandatory=$false)][string]$Country,
    [parameter(mandatory=$true)][string]$LicenseSku,
    [switch]$AddLicense,
    [switch]$RemoveLicense,
    $Credential

    )

<#
  .SYNOPSIS
  Add or Remove an O365 license to an account in the cloud 

  .DESCRIPTION
  Adds a license and usage location by using the -AddLicense parameter.
  Removes a license by using the -RemoveLicense parameter
  MSOL Service requires a UserPrincipalName input, so be sure to enter a UPN and not a samAccountName
  
  .EXAMPLE
    Add the specified license to the specified user
    Set-Office365License -UserPrincipalName jdoe@fabrikam.com -Country US -AddLicense -LicenseSku 'fabrikam:ENTERPRISE'

  .EXAMPLE
    Remove the specified license from the Msol user account.  Connects to Msol service with a prebuilt Credential object
    Set-Office365License -UserPrincipalName jdoe@fabrikam.com -RemoveLicense -LicenseSku 'fabrikam:ENTERPRISE' -Credential $Credential

  .SYNTAX
  Set-Office365License [-UserPrincipalName] <string> -LicenseSku <string> [-Country <string>] [-AddLicense] [-RemoveLicense] [-Credential <Object>]  [<CommonParameters>]

  .PARAMETER UserPrincipalName
  The Active Directory UserPrincipalName for the user

  .PARAMETER Country
  The ISO 3166-1 Alpha-2 two-letter country code where the licensed user resides

  .PARAMETER LicenseSku
  The MsolAccountSkuId for the license to add or remove.  Typically in 'domain:LICENSENAME' format

  .PARAMETER AddLicense
  Specifies that a license will be added to the user's Msol account

  .PARAMETER RemoveLicense
  Specifies that a license will be removed from the user's Msol account

  .PARAMETER Credential
  Credential object used to connect to the Msol Service with user account management permissions
  #>    

    BEGIN{
        
        if (($AddLicense -ne $true) -and ($RemoveLicense -ne $true)){
            $AddLicense = $true
        }

        if (!($Credential)){
            $Credential = Get-Credential -Message 'Enter O365 Admin Credentials'
        }

        Try #Connect to MSOL Server if not connected
        {
            Get-MsolDomain -ErrorAction Stop > $null
        }
        Catch 
        {
            Write-Host "Connecting to Office 365..."
            Connect-MsolService -Credential $Credential
        }       
                    
    }

    PROCESS{

        if ($AddLicense){
        
            if (!($Country)){

                $Country = Read-Host -Prompt "Please specify the two-digit country code"

            }
            
            Try{
                
                Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation $Country

                Try{
                    Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $LicenseSku -ErrorAction Stop
                }
                Catch{

                }

            }
            Catch{
                $_.exception.message
            
            }
        }

        elseif ($RemoveLicense){

            Try{
                Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $LicenseSku -ErrorAction Stop
            }
            Catch{
                $_.exception.message
            
            }

        }

        else{

        }
    
    }
    END{
        $checkmsol = Get-MsolUser -UserPrincipalName $UserPrincipalName
        
        if (($checkmsol.isLicensed -eq $true) -and ($AddLicense -eq $true)){
            Write-Host 'Office 365 account ' $UserPrincipalName ' has been provisioned successfully'  -ForegroundColor Green
            }
        elseif (($checkmsol.isLicensed -eq $False) -and ($AddLicense -eq $true)){
            Write-Host 'Office 365 account provisioning for ' $UserPrincipalName ' has failed' -ForegroundColor Red
            }      
        elseif (($checkmsol.isLicensed -eq $false) -and ($RemoveLicense -eq $true)){
            Write-Host 'Removal of Office 365 license for ' $UserPrincipalName ' has completed successfully'  -ForegroundColor Green
           }
        elseif(($checkmsol.isLicensed -eq $true) -and ($RemoveLicense -eq $true)){
            Write-Host 'Removal of Office 365 license for ' $UserPrincipalName ' has failed' -ForegroundColor Red
            } 
    }  

}