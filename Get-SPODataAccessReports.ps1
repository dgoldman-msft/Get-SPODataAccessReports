function Get-SPODataAccessReports {
    <#
        .SYNOPSIS
            This function exports all SPODataAccessReports.

        .DESCRIPTION
            Export all SPODataAccessReports.

        .PARAMETER DisconnectFromSPO
            Allow the user to specify whether to disconnect from the SPOService.

        .PARAMETER ReportEntity
            Specifies the entity that could cause oversharing and hence tracked by these reports.
            - EveryoneExceptExternalUsersAtSite
            - EveryoneExceptExternalUsersForItems
            - SharingLinks_Anyone
            - SharingLinks_PeopleInYourOrg
            - SharingLinks_Guests
            - SensitivityLabelForFiles
            - PermissionedUsers

        .PARAMETER TenantDomain
            The domain of the tenant.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso

            Command without verbose output

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -Verbose

            Command without verbose output

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity EveryoneExceptExternalUsersAtSite -ReportType Snapshot

            Export a report entity of EveryoneExceptExternalUsersAtSite and report type of RecentActivity

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -DoNotDisconnectFromSPO

            Allows the user to specify whether to disconnect from the SPOService or retain the current connection.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportType EveryoneExceptExternalUsersAtSite

            Selects a specific report type to export. (Default is all reports)

        .NOTES
            For more information please see: https://learn.microsoft.com/en-us/sharepoint/powershell-for-data-access-governance#creating-reports-using-powershell
    #>

    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Disconnect from SharePoint Online after the report collection is completed. Default is $false.')]
        [switch]
        $DisconnectFromSPO,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the entity that could cause oversharing and hence tracked by these reports. Valid values are: EveryoneExceptExternalUsersAtSite, EveryoneExceptExternalUsersForItems, SharingLinks_Anyone, SharingLinks_PeopleInYourOrg, SharingLinks_Guests, SensitivityLabelForFiles, PermissionedUsers.')]
        [ValidateSet('All', 'EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')]
        [string]
        $ReportEntity,

        [Parameter(Mandatory = $true, ParameterSetName = 'Default', HelpMessage = 'Specifies the domain of the tenant. This parameter is mandatory.')]
        [Parameter(Mandatory = $true)]
        [string]
        $TenantDomain,

        [Parameter(ParameterSetName = 'Default')]
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
        if ($ReportEntity -eq 'All') {
            $reportEntities = @('EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')
        }
        else {
            $reportEntities = @($ReportEntity)
        }

        # Iterate through each ReportEntity
        foreach ($entity in $reportEntities) {
            # Get the report data
            Write-Output "Getting report status for $($entity)."
            $reports = Get-SPODataAccessGovernanceInsight -ReportEntity $entity
            $reportArray = @()
            foreach ($report in $reports) {
                $reportArray += [PSCustomObject]@{
                    RunspaceId        = $report.RunspaceId
                    ReportId          = $report.ReportId
                    ReportName        = $report.ReportName
                    ReportEntity      = $report.ReportEntity
                    Status            = $report.Status
                    Workload          = $report.Workload
                    TriggeredDateTime = $report.TriggeredDateTime
                    CreatedDateTime   = $report.CreatedDateTime
                    ReportStartTime   = $report.ReportStartTime
                    ReportEndTime     = $report.ReportEndTime
                    ReportType        = $report.ReportType
                    SitesFound        = $report.SitesFound
                    Privacy           = $report.Privacy
                    Sensitivity       = $report.Sensitivity
                    Templates         = $report.Templates
                }
            }

            foreach ($report in $reportArray) {
                if ($report.Status -eq "Snapshot") {
                    Write-Output "NOTE: A 'Snapshot' report will have the latest data as of the report generation time and a 'RecentActivity' report will be based on data in the last 28 days."
                }

                switch ($report.Status) {
                    "NotStarted" {
                        Write-Output "Report generation for $($entity) has not yet begun."
                    }
                    "InQueue" {
                        Write-Output "Report for $($entity) is in the queue and waiting to be processed."
                    }
                    "InProgress" {
                        Write-Output "Report generation for $($entity) is currently in progress."
                    }
                    "Completed" {
                        Write-Output "Exporting completed report: $($entity)"
                        $report
                        Export-SPODataAccessGovernanceInsight -ReportID $report.ReportId
                    }
                    "Failed" {
                        Write-Output "Report generation for $($entity) has failed."
                    }
                    default {
                        Write-Output "Unknown report status: $($report.Status)"
                    }
                }
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
