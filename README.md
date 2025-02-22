# Get-SPODataAccessReports

Export all SPO Data Access Reports

## DESCRIPTION

This script will connect to your tenant and download all SPO Data Access Reports. Module can use PowerShell 5 and 7 to connect.

## EXAMPLE 1
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso

    Command without verbose output

## EXAMPLE 2
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -Verbose

    Command without verbose output

## EXAMPLE 3
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -DisconnectFromSPO

    Allows the user to specify whether to disconnect from the SPOService or retain the current connection.

## EXAMPLE 4
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportType EveryoneExceptExternalUsersAtSite

            Selects a specific report type to export. (Default is all reports)

## NOTES
- This command by default will NOT disconnect the SPO session when the script finishes. If you want to terminate your connect please see Example 3.

- For more information please see: https://learn.microsoft.com/en-us/sharepoint/powershell-for-data-access-governance#creating-reports-using-powershell

- This module needs to be ran as an Administrator in Powershell to export the reports.
