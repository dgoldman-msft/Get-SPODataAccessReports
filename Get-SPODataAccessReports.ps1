function Get-SPODataAccessReports {
    <#
        .SYNOPSIS
            This function exports all SPODataAccessReports.

        .DESCRIPTION
            Export all SPODataAccessReports.

        .PARAMETER DoNotDisconnectFromSPO
            Allow the user to specify whether to disconnect from the SPOService.

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

        .NOTES
            For more information please see: https://learn.microsoft.com/en-us/sharepoint/powershell-for-data-access-governance#creating-reports-using-powershell
    #>

    param (
        [CmdletBinding()]

        [switch]
        $DoNotDisconnectFromSPO,

        [Parameter(Mandatory = $true)]
        [string]
        $TenantDomain,

        [string]
        $TenantAdminUrl = "https://$TenantDomain-admin.sharepoint.com"
    )

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
            Import-Module -Name Microsoft.Online.SharePoint.PowerShell
            Write-Verbose "Imported Microsoft.Online.SharePoint.PowerShell module."
        }
        else {
            Write-Verbose "Microsoft.Online.SharePoint.PowerShell module already imported."
        }
    }
    catch {
        Write-Output "$_"
    }

    # Check connection to SharePoint Online
    try {
        Get-SPOTenant
        Write-Verbose "We are already connected to SharePoint Online."
    }
    catch {
        Write-Output "Not connected to SharePoint Online."
        try {
            # Connect to SharePoint Online
            Write-Verbose "Attempting to connect to SharePoint Online"
            Connect-SPOService -Url $TenantAdminUrl -ErrorAction SilentlyContinue
            Write-Output "Connected to SharePoint Online."
        }
        catch {
            write-Output "$_"
            return
        }
    }

    # Enumerate through all the ReportEntity values
    try {
        $reportEntities = @(
            [Microsoft.Online.SharePoint.TenantAdministration.ReportEntityEnum]::SharingLinks_Anyone,
            [Microsoft.Online.SharePoint.TenantAdministration.ReportEntityEnum]::SharingLinks_PeopleInYourOrg,
            [Microsoft.Online.SharePoint.TenantAdministration.ReportEntityEnum]::SharingLinks_Guests,
            [Microsoft.Online.SharePoint.TenantAdministration.ReportEntityEnum]::SensitivityLabelForFiles,
            [Microsoft.Online.SharePoint.TenantAdministration.ReportEntityEnum]::EveryoneExceptExternalUsersAtSite,
            [Microsoft.Online.SharePoint.TenantAdministration.ReportEntityEnum]::EveryoneExceptExternalUsersForItems,
            [Microsoft.Online.SharePoint.TenantAdministration.ReportEntityEnum]::PermissionedUsers
        )

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
        if ($DoNotDisconnectFromSPO -eq $True) {
            Write-Output "Disconnecting from the SPOService."
            Disconnect-SPOService
        }
        else {
            Write-Output "Not disconnecting from the SPOService."
        }
    }
}
