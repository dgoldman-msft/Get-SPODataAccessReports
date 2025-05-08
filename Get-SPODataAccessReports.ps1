function Write-ToLog {
    <#
        .SYNOPSIS
            Save output

        .DESCRIPTION
            Overload function for Write-Output

        .PARAMETER LoggingDirectory
            Directory to save the log file to. Default is "$env:MyDocuments".

        .PARAMETER LoggingFilename
            Filename to save the log file to. Default is "SamReportingLogs.txt".

        .EXAMPLE
            None

        .NOTES
            None
    #>

    [OutputType('System.String')]
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param
    (
        [Parameter(ParameterSetName = 'Default')]
        [string]
        $LoggingDirectory,

        [string]
        $LoggingFilename,

        [Parameter(Mandatory = $True, Position = 0)]
        [string]
        $InputString
    )

    try {
        if (-NOT(Test-Path -Path $LoggingDirectory)) {
            Write-Verbose "Creating New Logging Directory"
            New-Item -Path $LoggingDirectory -ItemType Directory -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Output "$_"
        return
    }

    try {
        # Console and log file output
        $stringObject = "[{0:MM/dd/yy} {0:HH:mm:ss}] - {1}" -f (Get-Date), $InputString
        Add-Content -Path (Join-Path $LoggingDirectory -ChildPath $LoggingFilename) -Value $stringObject -Encoding utf8 -ErrorAction Stop
    }
    catch {
        Write-Output "$_"
        return
    }
}

function Get-SPODataAccessReports {
    <#
        .SYNOPSIS
            This function exports all SPODataAccessReports.

        .DESCRIPTION
            Export all SPODataAccessReports.

        .PARAMETER DisconnectFromSPO
            Allow the user to specify whether to disconnect from the SPOService.

        .PARAMETER ExportReports
            Export reports to the specified directory. Default is "MyDocuments\Logging". If this parameter is not specified, the reports will not be exported.

        .PARAMETER LoggingDirectory
            Directory to save the log file to. Default is "MyDocuments\Logging".

        .PARAMETER LoggingFilename
            Filename to save the log file to. Default is "GetSamReportingLogs.txt".

        .PARAMETER ReportEntity
            Specifies the entity that could cause oversharing and hence tracked by these reports.
            - EveryoneExceptExternalUsersAtSite
            - EveryoneExceptExternalUsersForItems
            - SharingLinks_Anyone
            - SharingLinks_PeopleInYourOrg
            - SharingLinks_Guests
            - SensitivityLabelForFiles
            - PermissionedUsers

        .PARAMETER TableView
            Specifies the output report type. Valid outputs are: Filtered or Full.

        .PARAMETER TenantDomain
            The domain of the tenant.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All

            Get all reports

        .EXAMPLE
            C:\PS> Gsdar -TenantDomain Contoso -ReportEntity All

            Get all reports using alias. This will also work for all commands in this module.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All -TableView

            Get all reports (Default is all reports), when finished show all reports found in a table view

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All -ExportReports

            Export all reports to the specified directory. Default is "MyDocuments\Logging". If this parameter is not specified, the reports will not be exported.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All -ExportReports -SecondsToWait 120

            This sets a sleep timer so reports can be exported fully before renamed and moved to "MyDocuments\Logging". Default is 120 seconds. Maximum is 300 seconds.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All -ExportReports -SecondsToWait 120

            Export all reports to the specified directory. Default is "MyDocuments\Logging". If this parameter is not specified, the reports will not be exported.


        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -DoNotDisconnectFromSPO

            Allows the user to specify whether to disconnect from the SPOService or retain the current connection.

        .EXAMPLE
            C:\PS> Get-SPODataAccessReports -TenantDomain Contoso -ReportEntity All

            Get all reports. (Default is all reports)

        .NOTES
            For more information please see: https://learn.microsoft.com/en-us/sharepoint/powershell-for-data-access-governance#creating-reports-using-powershell
    #>

    [OutputType([System.Object])]
    [OutputType([System.String])]
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Alias('gsdar')]
        param (
        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Disconnect from SharePoint Online after the report collection is completed. Default is $false.')]
        [switch]
        $DisconnectFromSPO,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Export reports to the specified directory. Default is $env:MyDocuments\SamReporting.')]
        [switch]
        $ExportReports,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the directory to save the log file to. Default is $env:MyDocuments\SamReporting.')]
        [string]
        $LoggingDirectory = (Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "SamReporting"),

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the filename to save the log file to. Default is GetSamReportingLogs.txt.')]
        [string]
        $LoggingFilename = "GetSamReportingLogs.txt",

        [Parameter(Mandatory = $True, ParameterSetName = 'Default', HelpMessage = 'Specifies the entity that could cause oversharing and hence tracked by these reports. Valid values are: EveryoneExceptExternalUsersAtSite, EveryoneExceptExternalUsersForItems, SharingLinks_Anyone, SharingLinks_PeopleInYourOrg, SharingLinks_Guests, SensitivityLabelForFiles, PermissionedUsers.')]
        [ValidateSet('All', 'EveryoneExceptExternalUsersAtSite', 'EveryoneExceptExternalUsersForItems', 'SharingLinks_Anyone', 'SharingLinks_PeopleInYourOrg', 'SharingLinks_Guests', 'SensitivityLabelForFiles', 'PermissionedUsers')]
        [string]
        $ReportEntity = 'All',

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the number of seconds to wait for report export.')]
        [ValidateRange(1, 300)]
        [Int32]
        $SecondsToWait = 5,

        [Parameter(ParameterSetName = 'Default', HelpMessage = 'Specifies the output report type. Valid outputs are: Filtered or Full.')]
        [switch]
        $TableView,

        [Parameter(Mandatory = $true, ParameterSetName = 'Default', HelpMessage = 'Specifies the domain of the tenant. This parameter is mandatory.')]
        [string]
        $TenantDomain,

        [Parameter(ParameterSetName = 'Default')]
        [string]
        $TenantAdminUrl = "https://$TenantDomain-admin.sharepoint.com"
    )

    # Counters for report status
    $notStarted = 0
    $inQueue = 0
    $inProgress = 0
    $completed = 0
    $failed = 0
    $unknown = 0

    # Check if running as administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Output "This script must be run as an administrator."
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "This script must be run as an administrator."
        return
    }
    else {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Starting script execution as administrator."
    }

    try {
        # Check if the SharePoint Online Management Shell module is installed
        if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
            # Install the SharePoint Online Management Shell module
            Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber
            Write-Verbose "Installed Microsoft.Online.SharePoint.PowerShell module."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Installed Microsoft.Online.SharePoint.PowerShell module."
        }
        else {
            Write-Verbose "Microsoft.Online.SharePoint.PowerShell module already installed."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Microsoft.Online.SharePoint.PowerShell module already installed."
        }

        # Import the SharePoint Online Management Shell module
        if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell)) {
            if ($PSVersionTable.PSEdition -eq "Core") {
                Import-Module -Name Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -ErrorAction SilentlyContinue
                Write-Verbose "Connecting with Windows PowerShell Core Version."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting with Windows PowerShell Core Version."
            }
            else {
                Import-Module -Name Microsoft.Online.SharePoint.PowerShell
                Write-Verbose "Connecting with Windows PowerShell Desktop Version."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting with Windows PowerShell Desktop Version."
            }
        }
        else {
            Write-Verbose "Microsoft.Online.SharePoint.PowerShell module already imported."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Microsoft.Online.SharePoint.PowerShell already imported."
        }
    }
    catch {
        Write-Output "$_"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "ERROR: $_"
        return
    }

    # Check connection to SharePoint Online
    try {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connecting to SharePoint Online."
        Connect-SPOService -Url $TenantAdminUrl -ErrorAction SilentlyContinue
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Connected to SharePoint Online."
    }
    catch {
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Error: $_"
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
        $reportArray = @()
        foreach ($entity in $reportEntities) {
            # Get the report data
            $reports = Get-SPODataAccessGovernanceInsight -ReportEntity $entity
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
        }

        foreach ($report in $reportArray) {
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Getting report status for $($report.ReportEntity) - $($report.ReportId) - $($report.Status)"

            if ($report.Status -eq "Snapshot") {
                Write-Output "NOTE: A 'Snapshot' report will have the latest data as of the report generation time and a 'RecentActivity' report will be based on data in the last 28 days."
                Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "NOTE: A 'Snapshot' report will have the latest data as of the report generation time and a 'RecentActivity' report will be based on data in the last 28 days."
            }

            switch ($report.Status) {
                "NotStarted" {
                    Write-Output "Report generation for $($report.ReportId) has not yet begun."
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report generation for $($report.ReportEntity) - $($report.ReportId) has not yet begun."
                    $notStarted++
                }
                "InQueue" {
                    Write-Output "Report for $($report.ReportId) is in the queue and waiting to be processed."
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report for $($report.ReportEntity) - $($report.ReportId) is in the queue and waiting to be processed."
                    $inQueue++
                }
                "InProgress" {
                    Write-Output "Report generation for $($report.ReportId) is currently in progress."
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report generation for $($report.ReportEntity) - $($report.ReportId) is currently in progress."
                    $inProgress++
                }
                "Completed" {
                    # Temp fix until the Sharepoint Online Management Shell module is updated to reflect the DownloadPath parameter

                    if ($ExportReports) {
                        Export-SPODataAccessGovernanceInsight -ReportID $report.ReportId
                        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Exporting $($report.ReportEntity) - $($report.ReportId) completed!"
                        $dateString = (Get-Date).ToString("MMddyyyy_HHmmss")
                        Start-Sleep -Seconds $SecondsToWait
                        $exportPath = Get-ChildItem -Path . -Filter "*$($report.ReportId)*.csv" | Select-Object -First 1 -ExpandProperty FullName
                        if ($null -ne $exportPath) {
                            $newFileName = "$($report.ReportEntity)_$($report.ReportId)_$dateString.csv"
                            Rename-Item -Path $exportPath -NewName $newFileName
                            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Exported file for ReportId $($report.ReportId) renamed to $($newFileName)"
                            Move-Item -Path (Join-Path -Path (Get-Location) -ChildPath $newFileName) -Destination $LoggingDirectory
                            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report $($newFileName) moved to $($LoggingDirectory)"
                        }
                        else {
                            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Exported file for ReportId $($report.ReportId) not found."
                        }
                        Write-Output "Report renamed to $($newFileName) and moved to $($LoggingDirectory)"
                    }
                    else {
                        #Get-SPODataAccessGovernanceInsight -ReportID $report.ReportId
                    }

                    $completed++
                }
                "Failed" {
                    Write-Output "Report generation for $($report.ReportEntity) - $($report.ReportId) has failed."
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Report generation for $($report.ReportEntity) - $($report.ReportId) has failed."
                    $failed++
                }
                default {
                    Write-Output "Unknown report status: $($report.Status) for $($report.ReportEntity) - $($report.ReportId)"
                    Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Unknown report status: $($report.ReportEntity) - $($report.ReportId)"
                    $unknown++
                }
            }
        }


        if ($TableView) { $reportArray | Format-Table ReportId, ReportEntity, Status, Workload, ReportType, ReportType, SitesFound }
        else { foreach ($report in $reportArray) { $report } }
    }
    catch {
        Write-Output "$_"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "ERROR: $_"
    }
    finally {
        # Disconnect from Security & Compliance Center if connected
        Write-Output "`r`n-----------------------------------------"
        if ($DisconnectFromSPO -eq $True) {
            Write-Output "Disconnecting from the SPOService."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Disconnecting from the SPOService."
            Disconnect-SPOService
        }
        else {
            Write-Output "Not disconnecting from the SPOService."
            Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Not disconnecting from the SPOService."
        }

        Write-Output "`r`nTotal reports generated that are not started: $($notStarted)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Total reports generated that are not started: $($notStarted)"
        Write-Output "Total reports generated that are in queue: $($inQueue)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Total reports generated that are in queue: $($inQueue)"
        Write-Output "Total reports generated that are in progress: $($inProgress)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Total reports generated that are in progress: $($inProgress)"
        Write-Output "Total reports generated that failed: $($failed)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Total reports generated that failed: $($failed)"
        Write-Output "Total reports generated that are unknown: $($unknown)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Total reports generated that are unknown: $($unknown)"
        Write-Output "Total reports generated that are completed: $($completed)"
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Total reports generated that are completed: $($completed)"
        Write-Output "`r`nFor more information please see the logging file: $($LoggingDirectory)\$($LoggingFilename)"
        Write-Output "Script completed."
        Write-ToLog -LoggingDirectory $LoggingDirectory -LoggingFilename $LoggingFilename -InputString "Script completed."
    }
}
