# Get-SPODataAccessReports

Export all SPO Data Access Reports

## DESCRIPTION

This script will connect to your tenant and download all SPO Data Access Reports.

## EXAMPLE 1
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso

    Command without verbose output

## EXAMPLE 2
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -Verbose

    Command without verbose output

## EXAMPLE 3
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -DoNotDisconnectFromSPO

    Allows the user to specify whether to disconnect from the SPOService or retain the current connection.

## NOTES
- This command by default will disconnect the SPO session when the script finishes. If you want to retain your connect please see Example 3.

- For more information please see: https://learn.microsoft.com/en-us/sharepoint/powershell-for-data-access-governance#creating-reports-using-powershell

