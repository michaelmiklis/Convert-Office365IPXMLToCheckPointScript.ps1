######################################################################
## (C) 2018 Michael Miklis (michaelmiklis.de)
##
##
## Filename:      Convert-Office365IPXMLToCheckPointScript.ps1
##
## Version:       1.0
##
## Release:       Final
##
## Requirements:  -none-
##
## Description:   Converts the Microsoft Office 365 IP-Range XML File 
##                into a CheckPoint import script
##                
##                See https://support.office.com/en-gb/article/office-
##                365-urls-and-ip-address-ranges-8548a211-3fe7-47cb-
##                abb1-355ea5aa88a2 for more details
##
## This script is provided 'AS-IS'.  The author does not provide
## any guarantee or warranty, stated or implied.  Use at your own
## risk. You are free to reproduce, copy & modify the code, but
## please give the author credit.
##
####################################################################

param (
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()][string]$Uri,
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()][string]$Type = "IPv4"
)


Set-PSDebug -Strict
Set-StrictMode -Version latest
  
 
function Convert-Office365IPXMLToCheckPointScript
{
    <#
        .SYNOPSIS
        Converts XML File into CheckPoint import script
  
        .DESCRIPTION
        The Set-MSOLLicenseToADGroupMembers CMDlet downloads the
        XML file with all Office 365 urls, IPs, and subnets and
        creates a CheckPoint import script
  
        .PARAMETER Uri
        Uri or URL to the Microsoft XML file
  
        .EXAMPLE
        Convert-Office365IPXMLToCheckPointScript -Uri "https://go.microsoft.com/fwlink/?LinkId=533185"
 
    #>

    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()][string]$Uri,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()][string]$Type = "IPv4"
    )

    # Download xml file
	[Microsoft.PowerShell.Commands.WebResponseObject] $WebRequest = Invoke-WebRequest -Uri $Uri

    # Check if web request was successful
	if ($WebRequest.StatusCode -ne 200)
	{
		throw "HTTP request to $URI was not successful - received status code: $($WebRequest.StatusCode) ($($WebRequest.StatusDescription))"
	}

    # Try to cast response to XML
    try
    {
        $MicrosoftRanges = [xml]$WebRequest.Content
    }
    catch
    {
        throw "Unable to convert downloaded content to XML structure. Please verify the Uri."
    }

    # Loop through each product in XML
    foreach ($product in $MicrosoftRanges.products.product)
    {


        $addressEntries = $product.addresslist | Where-Object { $_.Type -eq $Type }

        # Log warning if no entries were found for $Type in $product
        if ($addressEntries -eq $null)
        {
            Write-Warning "Unable to find $Type in XML for product $($product.name)"
        }

        else
        {    
            
            if ($($addressEntries | Get-Member).Name.Contains("address") -eq $true)
            {
                # Loop through each address entry
                foreach ($addressEntry in $addressEntries.address)
                {

                    if ($Type.toLower() -eq "ipv4")
                    {
                        # Extract netmask
                        $netmask = [int]$addressEntry.Split("/")[1]
                        $netmask = ([convert]::ToInt64(('1' * $netmask + '0' * (32 - $netmask)), 2)) 
                        $netmask = '{0}.{1}.{2}.{3}' -f ([math]::Truncate($netmask / 16777216)).ToString(), ([math]::Truncate(($netmask % 16777216) / 65536)).ToString(), ([math]::Truncate(($netmask % 65536)/256)).ToString(), ([math]::Truncate($netmask % 256)).ToString() 

                        # Extract ipaddress
                        $ipaddress = $addressEntry.Split("/")[0]

                        "create network n_office365_{0}_{1}"-f $product.name, $ipaddress
                        "modify network_objects n_office365_{0}_{1} ipaddr {1}"-f $product.name, $ipaddress
                        "modify network_objects n_office365_{0}_{1} netmask {2}" -f $product.name, $ipaddress, $netmask
                    }

                    elseif ($Type.toLower() -eq "ipv6")
                    {
                        throw "not implemented"
                    }

                    elseif ($Type.toLower() -eq "url")
                    {
                        throw "not implemented"
                    }

                    else
                    {
                        throw "No valid type was specified. Valid types are IPv4, IPv6 or url"
                    }
                }
            }
        }


    }
    
}

Convert-Office365IPXMLToCheckPointScript -Uri $Uri -Type $Type
