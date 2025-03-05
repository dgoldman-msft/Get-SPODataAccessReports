# Get-SPODataAccessReports

Export all SPO Data Access Reports

## DESCRIPTION

This script will connect to your tenant and download all SPO Data Access Reports. Module can use PowerShell 5 and 7 to connect.

## EXAMPLE 1
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All

    Get all reports

## EXAMPLE 2
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All -TableView

    Get all reports (Default is all reports), when finished show all reports found in a table view

## EXAMPLE 3
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All -ExportReports

    Export all reports to the specified directory. Default is "MyDocuments\Logging". If this parameter is not specified, the reports will not be exported.

## EXAMPLE 4
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -DoNotDisconnectFromSPO

    Allows the user to specify whether to disconnect from the SPOService or retain the current connection.

## EXAMPLE 5
    C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All

    Get all reports. (Default is all reports)

## GENERAL NOTES
- This command by default will NOT disconnect the SPO session when the script finishes. If you want to terminate your connect please see Example 4.

- For more information please see: https://learn.microsoft.com/en-us/sharepoint/powershell-for-data-access-governance#creating-reports-using-powershell

- This module needs to be ran as an Administrator in Powershell to export the reports.

## USAGE NOTES
-  ReportType specifies the time period of data of the reports to be fetched. (Default is 'RecentActivity')
- 'Snapshot' report will have the latest data as of the report generation time.
- 'RecentActivity' report will be based on data in the last 28 days.
- 'Workload' default is 'SharePoint'.