﻿function Get-SPODataAccessReports {
    <#
        .SYNOPSIS
            This function exports all SPODataAccessReports.

        .DESCRIPTION
            Export all SPODataAccessReports.

        .PARAMETER DisconnectFromSPO
            Allow the user to specify whether to disconnect from the SPOService.

        .PARAMETER ReportType
            Specify the type of report to export. If not specified, all reports will be exported.

        .PARAMETER TenantDomain
            The domain of the tenant.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso

            Command without verbose output

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -Verbose

            Command without verbose output

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -DoNotDisconnectFromSPO

            Allows the user to specify whether to disconnect from the SPOService or retain the current connection.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportType EveryoneExceptExternalUsersAtSite

            Selects a specific report type to export. (Default is all reports)

        .NOTES
            For more information please see: https://learn.microsoft.com/en-us/sharepoint/powershell-for-data-access-governance#creating-reports-using-powershell
    #>

    param (
        [CmdletBinding(DefaultParameterSetName = 'Default')]
        [switch]
        $DisconnectFromSPO,

        [ValidateSet('EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')]
        [string]
        $ReportType,

        [Parameter(Mandatory = $true)]
        [string]
        $TenantDomain,

        [string]
        $TenantAdminUrl = "https://$TenantDomain-admin.sharepoint.com"
    )
    
    # Check if running as administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Output "This script must be run as an administrator."
        return
    }

    try {
        # Check if the SharePoint Online Management Shell module is installed
        if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
            # Install the SharePoint Online Management Shell module
            Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber
            Write-Verbose "Installed Microsoft.Online.SharePoint.PowerShell module."
        }
        else {
            Write-Verbose "Microsoft.Online.SharePoint.PowerShell module already installed."
        }

        # Import the SharePoint Online Management Shell module
        if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell)) {
            if ($PSVersionTable.PSEdition -eq "Core") {
                Import-Module -Name Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -ErrorAction SilentlyContinue
                Write-Verbose "Connecting with Windows PowerShell Core Version."
            }
            else {
                Import-Module -Name Microsoft.Online.SharePoint.PowerShell
                Write-Verbose "Connecting with Windows PowerShell Desktop Version."
            }
        }
        else {
            Write-Verbose "Microsoft.Online.SharePoint.PowerShell module already imported."
        }
    }
    catch {
        Write-Output "$_"
        return
    }

    # Check connection to SharePoint Online
    try {
        $connection = Get-SPOTenant -ErrorAction SilentlyContinue
        if (-not $connection) {
            Write-Output "Not connected to SharePoint Online. Attempting to connect to SharePoint Online"
            Connect-SPOService -Url $TenantAdminUrl -ErrorAction SilentlyContinue
            Write-Output "Connected to SharePoint Online."
        }
        else {
            Write-Verbose "Connected to SharePoint Online."
        }
    }
    catch {
        write-Output "$_"
        return
    }
    # Enumerate through all the ReportEntity values
    try {
        if($ReportType) {
            $reportEntities = @($ReportType)
        }
        else {
            $reportEntities = @('EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')
        }

        # Iterate through each ReportEntity
        foreach ($entity in $reportEntities) {
            # Get the report data
            Write-Output "Getting report status for $($entity)"
            $reports = Get-SPODataAccessGovernanceInsight -ReportEntity $entity

            # Check if there are any reports
            if ($reports.Status -eq "Completed") {
                # Iterate through each report and export it
                Write-Output "Exporting report: $($entity)"
                Export-SPODataAccessGovernanceInsight -ReportID $reports.ReportId
            }
        }
    }
    catch {
        Write-Output "$_"
    }
    finally {
        if ($DisconnectFromSPO -eq $True) {
            Write-Output "Disconnecting from the SPOService."
            Disconnect-SPOService
        }
        else {
            Write-Output "Not disconnecting from the SPOService."
        }
    }
}
