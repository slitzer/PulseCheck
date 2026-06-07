param(
    [string]$AppName = 'ITTool',
    [string]$RootPath = 'C:\Temp\ITTool',
    [string]$Ticket = '',
    [string]$Tech = '',
    [string]$Issue = ''
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$Script:Version = '0.1.0'
$Script:IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$Script:RunStarted = Get-Date
$Script:ComputerName = $env:COMPUTERNAME
$Script:CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$Script:RunName = '{0}_{1}' -f $Script:ComputerName, $Script:RunStarted.ToString('yyyy-MM-dd_HHmmss')
$Script:RunPath = Join-Path (Join-Path $RootPath 'Runs') $Script:RunName
$Script:LogPath = Join-Path $Script:RunPath 'Logs'
$Script:ReportAssetPath = Join-Path $Script:RunPath 'Reports'
$Script:ExportPath = Join-Path $Script:RunPath 'Exports'
$Script:TaskLogPath = Join-Path $Script:RunPath 'task-log.jsonl'
$Script:SummaryPath = Join-Path $Script:RunPath 'summary.json'
$Script:ReportPath = Join-Path $Script:RunPath 'report.html'
$Script:TranscriptPath = Join-Path $Script:RunPath 'transcript.log'
$Script:TaskResults = New-Object System.Collections.ArrayList
$Script:RecommendedTaskIds = New-Object System.Collections.Generic.HashSet[string]
$Script:Ui = @{}
$Script:TaskCheckBoxes = @{}
$Script:TranscriptStarted = $false

foreach ($path in @($RootPath, (Join-Path $RootPath 'Runs'), $Script:RunPath, $Script:LogPath, $Script:ReportAssetPath, $Script:ExportPath)) {
    if (-not (Test-Path -LiteralPath $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

try {
    Start-Transcript -Path $Script:TranscriptPath -Force | Out-Null
    $Script:TranscriptStarted = $true
} catch {
    Write-Warning "Unable to start transcript: $($_.Exception.Message)"
}

function ConvertTo-ItJson {
    param([Parameter(Mandatory = $true)]$InputObject, [int]$Depth = 8)
    $InputObject | ConvertTo-Json -Depth $Depth -Compress
}

function Write-TextFile {
    param([string]$Path, [AllowNull()]$Value)
    if ($null -eq $Value) { $Value = '' }
    [System.IO.File]::WriteAllText($Path, [string]$Value, [System.Text.Encoding]::UTF8)
}

function Export-CsvSafe {
    param([string]$Path, [AllowNull()]$Value)
    if ($null -eq $Value) {
        '' | Out-File -FilePath $Path -Encoding UTF8
        return
    }
    $Value | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Escape-Html {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Format-HtmlValue {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) { return (Escape-Html (($Value | ForEach-Object { [string]$_ }) -join ', ')) }
    return Escape-Html ([string]$Value)
}

function Get-SafePropertyValue {
    param([AllowNull()]$InputObject, [string]$Name)
    if ($null -eq $InputObject -or [string]::IsNullOrWhiteSpace($Name)) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    try {
        return $property.Value
    } catch {
        return ''
    }
}

function ConvertTo-HtmlTable {
    param([AllowNull()]$Rows, [int]$MaxRows = 300)
    if ($null -eq $Rows) { return '<p class="muted">No data collected.</p>' }
    $items = @($Rows)
    if (@($items).Count -eq 0) { return '<p class="muted">No data collected.</p>' }
    if ($items[0] -is [string]) {
        $text = ($items | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
        return '<pre>{0}</pre>' -f (Escape-Html $text)
    }
    $visibleItems = @($items | Select-Object -First $MaxRows)
    $props = @($items | ForEach-Object { $_.PSObject.Properties | Select-Object -ExpandProperty Name } | Select-Object -Unique)
    $head = ($props | ForEach-Object { '<th>{0}</th>' -f (Escape-Html $_) }) -join ''
    $body = foreach ($item in $visibleItems) {
        $cells = foreach ($prop in $props) {
            $value = Get-SafePropertyValue -InputObject $item -Name $prop
            if ($prop -eq 'Link' -and $value) {
                '<td><a href="{0}">{1}</a></td>' -f (Escape-Html ([string]$value)), (Escape-Html ([string]$value))
            } else {
                '<td>{0}</td>' -f (Format-HtmlValue $value)
            }
        }
        '<tr>{0}</tr>' -f ($cells -join '')
    }
    $itemCount = @($items).Count
    $note = if ($itemCount -gt $MaxRows) { '<p class="muted">Showing first {0} of {1} rows. Full data is available in the linked export or raw log.</p>' -f $MaxRows, $itemCount } else { '' }
    '<div class="table-wrap"><table><thead><tr>{0}</tr></thead><tbody>{1}</tbody></table></div>{2}' -f $head, ($body -join ''), $note
}

function Read-RelativeRunText {
    param([AllowNull()][string]$RelativePath, [int]$MaxChars = 12000)
    if ([string]::IsNullOrWhiteSpace($RelativePath)) { return '' }
    $path = Join-Path $Script:RunPath $RelativePath
    if (-not (Test-Path -LiteralPath $path)) { return '' }
    $text = Get-Content -LiteralPath $path -Raw -ErrorAction SilentlyContinue
    if ($null -eq $text) { return '' }
    if ($text.Length -gt $MaxChars) {
        return $text.Substring(0, $MaxChars) + "`r`n`r`n[Output truncated in report. Open the raw log for the full output.]"
    }
    return $text
}

function Convert-TaskDataToHtml {
    param($Task)
    $data = $Task.data
    if ($data -and @($data).Count -gt 0) {
        return ConvertTo-HtmlTable $data
    }

    $stdout = Read-RelativeRunText $Task.stdout
    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        return '<pre>{0}</pre>' -f (Escape-Html $stdout)
    }

    $stderr = Read-RelativeRunText $Task.stderr
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        return '<pre>{0}</pre>' -f (Escape-Html $stderr)
    }

    return '<p class="muted">No output captured.</p>'
}

function Get-RunFileRows {
    param([string]$Folder, [string]$Label)
    if (-not (Test-Path -LiteralPath $Folder)) { return @() }
    Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        ForEach-Object {
            [pscustomobject]@{
                Area = $Label
                File = $_.Name
                SizeKB = [math]::Round($_.Length / 1KB, 1)
                Link = $_.FullName.Substring($Script:RunPath.Length + 1)
            }
        }
}

function New-Task {
    param(
        [string]$Id,
        [string]$Name,
        [string]$Category,
        [bool]$AdminRequired,
        [bool]$Recommended,
        [scriptblock]$Action
    )
    if ($Recommended) { [void]$Script:RecommendedTaskIds.Add($Id) }
    [pscustomobject]@{
        Id = $Id
        Name = $Name
        Category = $Category
        AdminRequired = $AdminRequired
        Recommended = $Recommended
        Action = $Action
    }
}

function Get-InstalledApps {
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($path in $paths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation
    }
}

function Invoke-OptionalTextCommand {
    param([scriptblock]$Command)
    try {
        return (& $Command 2>$null | Out-String).Trim()
    } catch {
        return ''
    }
}

function Invoke-OptionalObjectCommand {
    param([scriptblock]$Command)
    try {
        return & $Command
    } catch {
        return $null
    }
}

function ConvertFrom-CimDate {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($text)
    } catch {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($text, [ref]$parsed)) {
            return $parsed
        }
        return $null
    }
}

function Get-MachineDetails {
    $cs = Invoke-OptionalObjectCommand { Get-CimInstance Win32_ComputerSystem -ErrorAction Stop }
    $os = Invoke-OptionalObjectCommand { Get-CimInstance Win32_OperatingSystem -ErrorAction Stop }
    $bios = Invoke-OptionalObjectCommand { Get-CimInstance Win32_BIOS -ErrorAction Stop }
    $ip = Invoke-OptionalObjectCommand { Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4Address } | Select-Object -First 1 }
    $battery = Invoke-OptionalObjectCommand { Get-CimInstance Win32_Battery -ErrorAction Stop | Select-Object -First 1 }
    $lastBoot = if ($os) { ConvertFrom-CimDate $os.LastBootUpTime } else { $null }
    $azureAdJoined = Invoke-OptionalTextCommand { dsregcmd /status } |
        Select-String 'AzureAdJoined' |
        ForEach-Object { ($_.Line -split ':', 2)[1].Trim() } |
        Select-Object -First 1
    [pscustomobject]@{
        Computer = $env:COMPUTERNAME
        User = $Script:CurrentUser
        Admin = $Script:IsAdmin
        DateTime = (Get-Date).ToString('s')
        OS = if ($os) { $os.Caption } else { '' }
        Build = if ($os) { $os.BuildNumber } else { '' }
        Uptime = if ($lastBoot) { ((Get-Date) - $lastBoot).ToString('d\.hh\:mm\:ss') } else { '' }
        LastBoot = $lastBoot
        Manufacturer = if ($cs) { $cs.Manufacturer } else { '' }
        Model = if ($cs) { $cs.Model } else { '' }
        SerialNumber = if ($bios) { $bios.SerialNumber } else { '' }
        DomainOrWorkgroup = if ($cs -and $cs.PartOfDomain) { $cs.Domain } elseif ($cs) { $cs.Workgroup } else { '' }
        AzureAdJoined = $azureAdJoined
        IPv4Address = if ($ip) { ($ip.IPv4Address.IPAddress -join ', ') } else { '' }
        Gateway = if ($ip -and $ip.IPv4DefaultGateway) { ($ip.IPv4DefaultGateway.NextHop -join ', ') } else { '' }
        DnsServers = if ($ip) { ($ip.DNSServer.ServerAddresses -join ', ') } else { '' }
        Battery = if ($battery) { '{0}% {1}' -f $battery.EstimatedChargeRemaining, $battery.BatteryStatus } else { 'Not detected' }
    }
}

function Get-ItTasks {
    @(
        New-Task 'details-machine-summary' 'Machine summary' 'Details' $false $true {
            $data = Get-MachineDetails
            Write-TextFile (Join-Path $Script:ExportPath 'machine-summary.json') (ConvertTo-ItJson $data)
            $data
        }
        New-Task 'details-os' 'OS details' 'Details' $false $true {
            Get-CimInstance Win32_OperatingSystem |
                Select-Object Caption, Version, BuildNumber, OSArchitecture, InstallDate, LastBootUpTime
        }
        New-Task 'details-uptime' 'Uptime / last boot' 'Details' $false $true {
            $os = Get-CimInstance Win32_OperatingSystem
            $lastBoot = ConvertFrom-CimDate $os.LastBootUpTime
            [pscustomobject]@{
                LastBoot = $lastBoot
                Uptime = if ($lastBoot) { ((Get-Date) - $lastBoot).ToString('d\.hh\:mm\:ss') } else { 'Unknown' }
            }
        }
        New-Task 'hardware-summary' 'CPU/RAM/GPU/BIOS' 'Hardware' $false $true {
            $ramGb = [math]::Round(((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB), 2)
            [pscustomobject]@{
                Cpu = ((Get-CimInstance Win32_Processor | Select-Object -First 1).Name)
                RamGB = $ramGb
                Gpu = ((Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name) -join '; ')
                Bios = ((Get-CimInstance Win32_BIOS | Select-Object -First 1).SMBIOSBIOSVersion)
            }
        }
        New-Task 'disk-logical' 'Logical disks and free space' 'Disk' $false $true {
            $data = Get-CimInstance Win32_LogicalDisk | Select-Object DeviceID, VolumeName,
                @{Name='SizeGB';Expression={[math]::Round($_.Size / 1GB, 2)}},
                @{Name='FreeGB';Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
                @{Name='FreePercent';Expression={if ($_.Size) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 2) } else { 0 }}}
            Export-CsvSafe (Join-Path $Script:ExportPath 'logical-disks.csv') $data
            $data
        }
        New-Task 'network-ipconfig' 'IP configuration' 'Network' $false $true { ipconfig /all }
        New-Task 'network-dns-client' 'DNS client config' 'Network' $false $true {
            Get-DnsClientServerAddress | Select-Object InterfaceAlias, AddressFamily, ServerAddresses
        }
        New-Task 'network-routes' 'Route table' 'Network' $false $true { route print }
        New-Task 'network-connectivity' 'Basic connectivity tests' 'Network' $false $true {
            [pscustomobject]@{ Target = '8.8.8.8'; Result = (Test-Connection 8.8.8.8 -Count 4 -Quiet) }
            [pscustomobject]@{ Target = '1.1.1.1'; Result = (Test-Connection 1.1.1.1 -Count 4 -Quiet) }
            Resolve-DnsName microsoft.com -ErrorAction SilentlyContinue | Select-Object -First 3
            Test-NetConnection microsoft.com -Port 443 -InformationLevel Detailed
        }
        New-Task 'network-wlan-report' 'WLAN report' 'Network' $false $true {
            netsh wlan show interfaces
            netsh wlan show wlanreport
            $source = Join-Path $env:ProgramData 'Microsoft\Windows\WlanReport\wlan-report-latest.html'
            if (Test-Path -LiteralPath $source) {
                Copy-Item -LiteralPath $source -Destination (Join-Path $Script:ReportAssetPath 'wlan-report-latest.html') -Force
            }
        }
        New-Task 'software-installed-apps' 'Installed applications from registry' 'Software' $false $true {
            $data = Get-InstalledApps | Sort-Object DisplayName -Unique
            Export-CsvSafe (Join-Path $Script:ExportPath 'installed-apps.csv') $data
            $data | Select-Object -First 250
        }
        New-Task 'software-startup-items' 'Startup items' 'Software' $false $true {
            $data = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location, User
            Export-CsvSafe (Join-Path $Script:ExportPath 'startup-items.csv') $data
            $data
        }
        New-Task 'software-processes' 'Running processes' 'Software' $false $true {
            $data = Get-Process |
                Select-Object Name, Id,
                    @{Name='CPUSeconds';Expression={ if ($_.CPU -is [timespan]) { [math]::Round($_.CPU.TotalSeconds, 2) } elseif ($null -ne $_.CPU) { [math]::Round([double]$_.CPU, 2) } else { $null } }},
                    WorkingSet,
                    @{Name='Path';Expression={ try { $_.Path } catch { '' } }} |
                Sort-Object CPUSeconds -Descending
            Export-CsvSafe (Join-Path $Script:ExportPath 'processes.csv') $data
            $data | Select-Object -First 100
        }
        New-Task 'services-status' 'Services and status' 'Services' $false $true {
            $data = Get-CimInstance Win32_Service | Select-Object Name, DisplayName, State, StartMode, StartName, PathName
            Export-CsvSafe (Join-Path $Script:ExportPath 'services.csv') $data
            $data
        }
        New-Task 'updates-hotfixes' 'Installed hotfixes' 'Updates' $false $true {
            $data = Get-HotFix | Select-Object HotFixID, Description, InstalledBy, InstalledOn
            Export-CsvSafe (Join-Path $Script:ExportPath 'hotfixes.csv') $data
            $data
        }
        New-Task 'events-system' 'System event log errors/warnings' 'Events' $false $true {
            $data = Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1,2,3; StartTime = (Get-Date).AddDays(-7) } -MaxEvents 500 |
                Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message
            Export-CsvSafe (Join-Path $Script:ExportPath 'eventlog-system.csv') $data
            $data | Select-Object -First 100
        }
        New-Task 'events-application' 'Application event log errors/warnings' 'Events' $false $true {
            $data = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Level = 1,2,3; StartTime = (Get-Date).AddDays(-7) } -MaxEvents 500 |
                Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message
            Export-CsvSafe (Join-Path $Script:ExportPath 'eventlog-application.csv') $data
            $data | Select-Object -First 100
        }
        New-Task 'reliability-records' 'Reliability monitor records' 'Reliability' $false $true {
            $data = Get-CimInstance -ClassName Win32_ReliabilityRecords -ErrorAction SilentlyContinue |
                Where-Object { $_.TimeGenerated -gt (Get-Date).AddDays(-7) } |
                Select-Object TimeGenerated, SourceName, ProductName, EventIdentifier, Message
            Export-CsvSafe (Join-Path $Script:ExportPath 'reliability-records.csv') $data
            $data
        }
        New-Task 'power-battery-report' 'Battery report' 'Power' $false $true {
            powercfg /batteryreport /output (Join-Path $Script:ReportAssetPath 'battery-report.html')
        }
        New-Task 'events-security' 'Security event log' 'Events' $true $false {
            Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Level = 1,2,3; StartTime = (Get-Date).AddDays(-7) } -MaxEvents 500 |
                Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message
        }
        New-Task 'health-sfc' 'sfc /scannow' 'Health' $true $false { sfc /scannow }
        New-Task 'health-dism-scanhealth' 'DISM ScanHealth' 'Health' $true $false { DISM /Online /Cleanup-Image /ScanHealth }
        New-Task 'power-energy-report' 'powercfg /energy' 'Power' $true $false {
            powercfg /energy /output (Join-Path $Script:ReportAssetPath 'energy-report.html')
        }
        New-Task 'drivers-export' 'Driver export' 'Drivers' $true $false {
            $driverPath = Join-Path $Script:ExportPath 'drivers'
            New-Item -ItemType Directory -Path $driverPath -Force | Out-Null
            Export-WindowsDriver -Online -Destination $driverPath
        }
    )
}

function Invoke-ItTask {
    param([Parameter(Mandatory = $true)]$Task, [int]$Index)
    $safeName = ($Task.Id -replace '[^a-zA-Z0-9\-]', '-').ToLowerInvariant()
    $prefix = '{0:000}-{1}' -f $Index, $safeName
    $stdoutRel = Join-Path 'Logs' ($prefix + '.out.txt')
    $stderrRel = Join-Path 'Logs' ($prefix + '.err.txt')
    $stdoutPath = Join-Path $Script:RunPath $stdoutRel
    $stderrPath = Join-Path $Script:RunPath $stderrRel
    $started = Get-Date
    $status = 'Success'
    $exitCode = $null
    $exception = $null
    $resultData = $null

    if ($Task.AdminRequired -and -not $Script:IsAdmin) {
        $ended = Get-Date
        $result = [pscustomobject]@{
            taskId = $Task.Id; name = $Task.Name; category = $Task.Category; adminRequired = $Task.AdminRequired
            run = $false; status = 'Skipped'; reason = 'Admin required but tool is not elevated'
            command = $Task.Action.ToString(); startTime = $started.ToString('o'); endTime = $ended.ToString('o')
            durationMs = 0; exitCode = $null; stdout = $stdoutRel; stderr = $stderrRel; exception = $null
        }
        Write-TextFile $stdoutPath ''
        Write-TextFile $stderrPath $result.reason
        Add-TaskResult $result $null
        return $result
    }

    try {
        $global:LASTEXITCODE = $null
        $output = & $Task.Action 2>&1
        $exitCode = if ($global:LASTEXITCODE -ne $null) { $global:LASTEXITCODE } else { 0 }
        $errors = @($output | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
        $normal = @($output | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] })
        Write-TextFile $stdoutPath ($normal | Out-String -Width 240)
        Write-TextFile $stderrPath ($errors | Out-String -Width 240)
        $resultData = $normal
        if ($errors.Count -gt 0 -or ($exitCode -is [int] -and $exitCode -ne 0)) {
            $status = 'Warning'
        }
    } catch {
        $status = 'Failed'
        $exception = $_.Exception.Message
        Write-TextFile $stdoutPath ''
        Write-TextFile $stderrPath ($_ | Out-String -Width 240)
    }

    $ended = Get-Date
    $result = [pscustomobject]@{
        taskId = $Task.Id; name = $Task.Name; category = $Task.Category; adminRequired = $Task.AdminRequired
        run = $true; status = $status; reason = $null; command = $Task.Action.ToString()
        startTime = $started.ToString('o'); endTime = $ended.ToString('o')
        durationMs = [int](New-TimeSpan -Start $started -End $ended).TotalMilliseconds
        exitCode = $exitCode; stdout = $stdoutRel; stderr = $stderrRel; exception = $exception
    }
    Add-TaskResult $result $resultData
    return $result
}

function Add-TaskResult {
    param($Result, [AllowNull()]$Data)
    $Result | Add-Member -NotePropertyName data -NotePropertyValue $Data -Force
    [void]$Script:TaskResults.Add($Result)
    $logObject = $Result | Select-Object taskId, name, category, adminRequired, run, status, reason, command, startTime, endTime, durationMs, exitCode, stdout, stderr, exception
    Add-Content -Path $Script:TaskLogPath -Value (ConvertTo-ItJson $logObject) -Encoding UTF8
}

function Save-Summary {
    $counts = $Script:TaskResults | Group-Object status | ForEach-Object { [pscustomobject]@{ Status = $_.Name; Count = $_.Count } }
    $summary = [pscustomobject]@{
        appName = $AppName
        version = $Script:Version
        runPath = $Script:RunPath
        reportPath = $Script:ReportPath
        computer = $Script:ComputerName
        user = $Script:CurrentUser
        admin = $Script:IsAdmin
        ticket = $Script:Ui.TicketBox.Text
        tech = $Script:Ui.TechBox.Text
        issue = $Script:Ui.IssueBox.Text
        started = $Script:RunStarted.ToString('o')
        finished = (Get-Date).ToString('o')
        counts = $counts
        tasks = ($Script:TaskResults | Select-Object taskId, name, category, adminRequired, run, status, reason, durationMs, exitCode, stdout, stderr, exception)
    }
    Write-TextFile $Script:SummaryPath (ConvertTo-Json $summary -Depth 8)
}

function New-Report {
    $finished = Get-Date
    $duration = New-TimeSpan -Start $Script:RunStarted -End $finished
    $taskCount = @($Script:TaskResults).Count
    $successCount = @($Script:TaskResults | Where-Object { $_.status -eq 'Success' }).Count
    $warningCount = @($Script:TaskResults | Where-Object { $_.status -eq 'Warning' }).Count
    $failedCount = @($Script:TaskResults | Where-Object { $_.status -eq 'Failed' }).Count
    $skippedCount = @($Script:TaskResults | Where-Object { $_.status -eq 'Skipped' }).Count
    $summaryRows = @(
        [pscustomobject]@{ Field = 'Computer'; Value = $Script:ComputerName },
        [pscustomobject]@{ Field = 'User'; Value = $Script:CurrentUser },
        [pscustomobject]@{ Field = 'Admin'; Value = $Script:IsAdmin },
        [pscustomobject]@{ Field = 'Ticket'; Value = $Script:Ui.TicketBox.Text },
        [pscustomobject]@{ Field = 'Tech'; Value = $Script:Ui.TechBox.Text },
        [pscustomobject]@{ Field = 'Issue'; Value = $Script:Ui.IssueBox.Text },
        [pscustomobject]@{ Field = 'Started'; Value = $Script:RunStarted },
        [pscustomobject]@{ Field = 'Finished'; Value = $finished },
        [pscustomobject]@{ Field = 'Duration'; Value = $duration.ToString('hh\:mm\:ss') },
        [pscustomobject]@{ Field = 'Run folder'; Value = $Script:RunPath }
    )
    $taskTable = ConvertTo-HtmlTable ($Script:TaskResults | Select-Object name, category, status, durationMs, stdout, stderr, reason, exception)
    $machineDetailRows = (Get-MachineDetails).PSObject.Properties | ForEach-Object { [pscustomobject]@{ Field = $_.Name; Value = $_.Value } }
    $machineDetails = ConvertTo-HtmlTable $machineDetailRows
    $files = @()
    $files += [pscustomobject]@{ Area = 'Run root'; File = 'report.html'; SizeKB = ''; Link = 'report.html' }
    $files += [pscustomobject]@{ Area = 'Run root'; File = 'summary.json'; SizeKB = ''; Link = 'summary.json' }
    $files += [pscustomobject]@{ Area = 'Run root'; File = 'task-log.jsonl'; SizeKB = ''; Link = 'task-log.jsonl' }
    $files += [pscustomobject]@{ Area = 'Run root'; File = 'transcript.log'; SizeKB = ''; Link = 'transcript.log' }
    $files += [pscustomobject]@{ Area = 'Reports'; File = 'full-report.html'; SizeKB = ''; Link = 'Reports\full-report.html' }
    $files += Get-RunFileRows -Folder $Script:LogPath -Label 'Logs'
    $files += Get-RunFileRows -Folder $Script:ExportPath -Label 'Exports'
    $files += Get-RunFileRows -Folder $Script:ReportAssetPath -Label 'Reports'
    $fileTable = ConvertTo-HtmlTable $files

    $navLinks = foreach ($group in ($Script:TaskResults | Group-Object category | Sort-Object Name)) {
        '<a href="#{0}">{1}</a>' -f (Escape-Html ($group.Name -replace '[^a-zA-Z0-9]+', '-').Trim('-').ToLowerInvariant()), (Escape-Html $group.Name)
    }
    $sections = foreach ($group in ($Script:TaskResults | Group-Object category | Sort-Object Name)) {
        $sectionId = ($group.Name -replace '[^a-zA-Z0-9]+', '-').Trim('-').ToLowerInvariant()
        $items = foreach ($task in $group.Group) {
            $badge = '<span class="badge {0}">{0}</span>' -f (Escape-Html $task.status.ToLowerInvariant())
            $raw = '<a href="{0}">stdout</a><a href="{1}">stderr</a>' -f (Escape-Html $task.stdout), (Escape-Html $task.stderr)
            $metaRows = @(
                [pscustomobject]@{ Field = 'Task ID'; Value = $task.taskId },
                [pscustomobject]@{ Field = 'Duration'; Value = ('{0} ms' -f $task.durationMs) },
                [pscustomobject]@{ Field = 'Exit code'; Value = $task.exitCode },
                [pscustomobject]@{ Field = 'Admin required'; Value = $task.adminRequired },
                [pscustomobject]@{ Field = 'Reason'; Value = $task.reason },
                [pscustomobject]@{ Field = 'Exception'; Value = $task.exception }
            )
            $dataHtml = Convert-TaskDataToHtml $task
            '<details class="task"><summary>{0}<strong>{1}</strong><span>{2} ms</span></summary><div class="task-body"><div class="task-links">{3}</div><h3>Result</h3>{4}<h3>Metadata</h3>{5}</div></details>' -f $badge, (Escape-Html $task.name), (Escape-Html $task.durationMs), $raw, $dataHtml, (ConvertTo-HtmlTable $metaRows)
        }
        '<section id="{0}"><h2>{1}</h2>{2}</section>' -f (Escape-Html $sectionId), (Escape-Html $group.Name), ($items -join "`n")
    }
    $failed = ConvertTo-HtmlTable ($Script:TaskResults | Where-Object { $_.status -eq 'Failed' -or $_.status -eq 'Skipped' } | Select-Object name, category, status, reason, exception)
    $html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$([System.Net.WebUtility]::HtmlEncode($AppName)) Report</title>
<style>
:root { color-scheme: light dark; --bg:#f4f6f8; --panel:#fff; --panel2:#f9fafb; --text:#17202a; --muted:#667085; --line:#d7dde5; --ok:#157347; --warn:#b7791f; --fail:#b42318; --skip:#475467; --accent:#0f766e; --accent2:#164e63; }
@media (prefers-color-scheme: dark) { :root { --bg:#101418; --panel:#171c22; --panel2:#11161c; --text:#e6edf3; --muted:#a7b0bc; --line:#303844; } }
* { box-sizing:border-box; }
body { margin:0; font-family: Segoe UI, Arial, sans-serif; background:var(--bg); color:var(--text); line-height:1.42; }
header { background:linear-gradient(135deg, var(--accent2), var(--accent)); color:#fff; padding:28px 36px; }
h1 { margin:0 0 6px; font-size:30px; font-weight:650; }
h2 { margin:0 0 14px; font-size:20px; }
h3 { margin:18px 0 8px; font-size:15px; color:var(--muted); text-transform:uppercase; }
main { padding:24px 32px 44px; max-width:1440px; margin:0 auto; }
nav { display:flex; flex-wrap:wrap; gap:8px; padding:14px 32px; background:var(--panel); border-bottom:1px solid var(--line); position:sticky; top:0; z-index:2; }
nav a, .task-links a { color:var(--accent); text-decoration:none; border:1px solid var(--line); border-radius:6px; padding:6px 9px; background:var(--panel); }
.subtle { opacity:.86; }
.grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:12px; margin-bottom:18px; }
.metric, section, details.task { background:var(--panel); border:1px solid var(--line); border-radius:8px; }
.metric { padding:14px 16px; }
.metric span { color:var(--muted); display:block; }
.metric strong { display:block; font-size:26px; margin-top:3px; }
section { padding:18px; margin-bottom:18px; }
details.task { margin:10px 0; overflow:hidden; }
summary { cursor:pointer; padding:12px 14px; display:flex; gap:10px; align-items:center; list-style:none; }
summary::-webkit-details-marker { display:none; }
summary strong { font-size:14px; }
summary span:last-child { margin-left:auto; color:var(--muted); }
.task-body { padding:0 14px 16px; border-top:1px solid var(--line); }
.task-links { display:flex; gap:8px; flex-wrap:wrap; margin:12px 0; }
.badge { font-size:11px; font-weight:700; padding:3px 8px; border-radius:999px; color:#fff; text-transform:uppercase; min-width:72px; text-align:center; }
.success { background:var(--ok); } .warning { background:var(--warn); } .failed { background:var(--fail); } .skipped { background:var(--skip); }
.muted { color:var(--muted); }
.table-wrap { overflow:auto; border:1px solid var(--line); border-radius:6px; background:var(--panel); }
table { border-collapse:collapse; min-width:100%; font-size:13px; }
th, td { padding:8px 10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; max-width:760px; overflow-wrap:anywhere; }
th { background:var(--panel2); color:var(--muted); font-weight:650; }
pre { white-space:pre-wrap; overflow:auto; max-height:520px; padding:12px; margin:0; border:1px solid var(--line); border-radius:6px; background:var(--panel2); font:12px Consolas, monospace; }
a { color:var(--accent); }
</style>
</head>
<body>
<header>
<h1>$([System.Net.WebUtility]::HtmlEncode($AppName)) Diagnostic Report</h1>
<div class="subtle">$([System.Net.WebUtility]::HtmlEncode($Script:ComputerName)) | $([System.Net.WebUtility]::HtmlEncode($finished.ToString('yyyy-MM-dd HH:mm:ss'))) | Run folder: $([System.Net.WebUtility]::HtmlEncode($Script:RunPath))</div>
</header>
<nav><a href="#summary">Summary</a><a href="#machine-details">Machine Details</a><a href="#task-results">Task Results</a><a href="#files">Files</a>$($navLinks -join '')</nav>
<main>
<div class="grid">
<div class="metric"><span>Tasks</span><strong>$taskCount</strong></div>
<div class="metric"><span>Success</span><strong>$successCount</strong></div>
<div class="metric"><span>Warnings</span><strong>$warningCount</strong></div>
<div class="metric"><span>Failed</span><strong>$failedCount</strong></div>
<div class="metric"><span>Skipped</span><strong>$skippedCount</strong></div>
</div>
<section id="summary"><h2>Summary</h2>$(ConvertTo-HtmlTable $summaryRows)</section>
<section id="machine-details"><h2>Machine Details</h2>$machineDetails</section>
<section id="task-results"><h2>Task Results</h2>$taskTable</section>
<section id="failed-skipped"><h2>Failed And Skipped Tasks</h2>$failed</section>
<section id="files"><h2>Generated Files</h2>$fileTable</section>
$($sections -join "`n")
</main>
</body>
</html>
"@
    Write-TextFile $Script:ReportPath $html
    Write-TextFile (Join-Path $Script:ReportAssetPath 'full-report.html') $html
}

function Update-UiStatus {
    param([string]$Text, [int]$Completed, [int]$Total)
    if ($Script:Ui.StatusText) { $Script:Ui.StatusText.Text = $Text }
    if ($Script:Ui.ProgressBar) {
        $Script:Ui.ProgressBar.Maximum = [Math]::Max($Total, 1)
        $Script:Ui.ProgressBar.Value = $Completed
    }
    if ($Script:Ui.ProgressText) { $Script:Ui.ProgressText.Text = '{0} / {1} tasks complete' -f $Completed, $Total }
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-SelectedTasks {
    param([bool]$RecommendedOnly)
    $tasks = Get-ItTasks
    $selected = foreach ($task in $tasks) {
        if ($RecommendedOnly) {
            if ($task.Recommended) { $task }
        } elseif ($Script:TaskCheckBoxes.ContainsKey($task.Id) -and $Script:TaskCheckBoxes[$task.Id].IsChecked) {
            $task
        }
    }
    $selected = @($selected)
    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No diagnostic tasks are selected.', $AppName, 'OK', 'Information') | Out-Null
        return
    }
    $Script:TaskResults.Clear()
    if (Test-Path -LiteralPath $Script:TaskLogPath) { Remove-Item -LiteralPath $Script:TaskLogPath -Force }
    $Script:Ui.RunSelectedButton.IsEnabled = $false
    $Script:Ui.RunRecommendedButton.IsEnabled = $false
    try {
        $i = 0
        foreach ($task in $selected) {
            $i++
            Update-UiStatus ('Collecting {0}...' -f $task.Name) ($i - 1) $selected.Count
            Invoke-ItTask -Task $task -Index $i | Out-Null
            Update-UiStatus ('Completed {0}' -f $task.Name) $i $selected.Count
        }
        try {
            Save-Summary
            New-Report
            Update-UiStatus 'Complete. Report generated.' $selected.Count $selected.Count
            $Script:Ui.OpenReportButton.IsEnabled = $true
            $Script:Ui.OpenFolderButton.IsEnabled = $true
            $Script:Ui.ZipButton.IsEnabled = $true
        } catch {
            $message = "Diagnostics completed, but report generation failed: $($_.Exception.Message)"
            Write-TextFile (Join-Path $Script:RunPath 'report-error.txt') ($_ | Out-String -Width 240)
            Update-UiStatus $message $selected.Count $selected.Count
            $Script:Ui.OpenFolderButton.IsEnabled = $true
            [System.Windows.MessageBox]::Show($message, $AppName, 'OK', 'Error') | Out-Null
        }
    } finally {
        $Script:Ui.RunSelectedButton.IsEnabled = $true
        $Script:Ui.RunRecommendedButton.IsEnabled = $true
    }
}

function New-Ui {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms

    $window = New-Object System.Windows.Window
    $window.Title = $AppName
    $window.Width = 1040
    $window.Height = 720
    $window.MinWidth = 860
    $window.MinHeight = 600
    $window.WindowStartupLocation = 'CenterScreen'
    $window.FontFamily = 'Segoe UI'
    $window.Background = '#f5f7fb'

    $root = New-Object System.Windows.Controls.DockPanel
    $root.Margin = '18'
    $window.Content = $root

    $header = New-Object System.Windows.Controls.StackPanel
    $header.Margin = '0,0,0,14'
    [System.Windows.Controls.DockPanel]::SetDock($header, 'Top')
    $root.Children.Add($header) | Out-Null

    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = $AppName
    $title.FontSize = 28
    $title.FontWeight = 'SemiBold'
    $header.Children.Add($title) | Out-Null

    $subtitle = New-Object System.Windows.Controls.TextBlock
    $subtitle.Text = 'Windows Diagnostic & Reporting Tool'
    $subtitle.Foreground = '#5d6675'
    $subtitle.Margin = '0,2,0,10'
    $header.Children.Add($subtitle) | Out-Null

    $identity = New-Object System.Windows.Controls.TextBlock
    $identity.Text = 'Computer: {0}    User: {1}    Admin: {2}' -f $Script:ComputerName, $Script:CurrentUser, ($(if ($Script:IsAdmin) { 'Yes' } else { 'No' }))
    $identity.FontWeight = 'SemiBold'
    $header.Children.Add($identity) | Out-Null

    $form = New-Object System.Windows.Controls.Grid
    $form.Margin = '0,12,0,0'
    0..5 | ForEach-Object {
        $col = New-Object System.Windows.Controls.ColumnDefinition
        $col.Width = if ($_ % 2 -eq 0) { 'Auto' } else { '*' }
        $form.ColumnDefinitions.Add($col)
    }
    foreach ($rowIndex in 0,1) {
        $row = New-Object System.Windows.Controls.RowDefinition
        $row.Height = 'Auto'
        $form.RowDefinitions.Add($row)
    }
    $header.Children.Add($form) | Out-Null

    function Add-LabelBox([string]$labelText, [int]$row, [int]$labelCol, [string]$value, [int]$span) {
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = $labelText
        $label.Margin = '0,4,8,4'
        $box = New-Object System.Windows.Controls.TextBox
        $box.Text = $value
        $box.Margin = '0,2,16,6'
        $box.MinWidth = 120
        [System.Windows.Controls.Grid]::SetRow($label, $row)
        [System.Windows.Controls.Grid]::SetColumn($label, $labelCol)
        [System.Windows.Controls.Grid]::SetRow($box, $row)
        [System.Windows.Controls.Grid]::SetColumn($box, $labelCol + 1)
        [System.Windows.Controls.Grid]::SetColumnSpan($box, $span)
        $form.Children.Add($label) | Out-Null
        $form.Children.Add($box) | Out-Null
        return $box
    }
    $Script:Ui.TicketBox = Add-LabelBox 'Ticket:' 0 0 $Ticket 1
    $Script:Ui.TechBox = Add-LabelBox 'Tech:' 0 2 $Tech 1
    $Script:Ui.IssueBox = Add-LabelBox 'Issue:' 1 0 $Issue 5

    $footer = New-Object System.Windows.Controls.StackPanel
    $footer.Margin = '0,14,0,0'
    [System.Windows.Controls.DockPanel]::SetDock($footer, 'Bottom')
    $root.Children.Add($footer) | Out-Null

    $progressGrid = New-Object System.Windows.Controls.Grid
    $progressGrid.Margin = '0,0,0,10'
    $progressColumn = New-Object System.Windows.Controls.ColumnDefinition
    $progressColumn.Width = New-Object System.Windows.GridLength 1, ([System.Windows.GridUnitType]::Star)
    $progressTextColumn = New-Object System.Windows.Controls.ColumnDefinition
    $progressTextColumn.Width = [System.Windows.GridLength]::Auto
    $progressGrid.ColumnDefinitions.Add($progressColumn)
    $progressGrid.ColumnDefinitions.Add($progressTextColumn)
    $footer.Children.Add($progressGrid) | Out-Null

    $Script:Ui.ProgressBar = New-Object System.Windows.Controls.ProgressBar
    $Script:Ui.ProgressBar.Height = 18
    [System.Windows.Controls.Grid]::SetColumn($Script:Ui.ProgressBar, 0)
    $progressGrid.Children.Add($Script:Ui.ProgressBar) | Out-Null

    $Script:Ui.ProgressText = New-Object System.Windows.Controls.TextBlock
    $Script:Ui.ProgressText.Margin = '12,0,0,0'
    $Script:Ui.ProgressText.Text = '0 / 0 tasks complete'
    [System.Windows.Controls.Grid]::SetColumn($Script:Ui.ProgressText, 1)
    $progressGrid.Children.Add($Script:Ui.ProgressText) | Out-Null

    $Script:Ui.StatusText = New-Object System.Windows.Controls.TextBlock
    $Script:Ui.StatusText.Text = 'Ready.'
    $Script:Ui.StatusText.Foreground = '#475467'
    $footer.Children.Add($Script:Ui.StatusText) | Out-Null

    $buttons = New-Object System.Windows.Controls.StackPanel
    $buttons.Orientation = 'Horizontal'
    $buttons.Margin = '0,10,0,0'
    $footer.Children.Add($buttons) | Out-Null

    function New-Button([string]$text) {
        $button = New-Object System.Windows.Controls.Button
        $button.Content = $text
        $button.Margin = '0,0,8,0'
        $button.Padding = '14,7'
        $button.MinWidth = 110
        $buttons.Children.Add($button) | Out-Null
        return $button
    }
    $Script:Ui.RunSelectedButton = New-Button 'Run Selected'
    $Script:Ui.RunRecommendedButton = New-Button 'Run Recommended'
    $Script:Ui.OpenReportButton = New-Button 'Open Report'
    $Script:Ui.OpenFolderButton = New-Button 'Open Folder'
    $Script:Ui.ZipButton = New-Button 'Create ZIP'
    $Script:Ui.OpenReportButton.IsEnabled = $false
    $Script:Ui.OpenFolderButton.IsEnabled = $true
    $Script:Ui.ZipButton.IsEnabled = $false

    $tabs = New-Object System.Windows.Controls.TabControl
    $root.Children.Add($tabs) | Out-Null

    function Add-Tab([string]$name, [System.Windows.UIElement]$content) {
        $tab = New-Object System.Windows.Controls.TabItem
        $tab.Header = $name
        $tab.Content = $content
        $tabs.Items.Add($tab) | Out-Null
    }

    $detailsScroll = New-Object System.Windows.Controls.ScrollViewer
    $detailsScroll.VerticalScrollBarVisibility = 'Auto'
    $detailsScroll.HorizontalScrollBarVisibility = 'Auto'
    $detailsPanel = New-Object System.Windows.Controls.StackPanel
    $detailsPanel.Margin = '12'
    $detailsScroll.Content = $detailsPanel
    $details = Get-MachineDetails
    foreach ($prop in $details.PSObject.Properties) {
        $line = New-Object System.Windows.Controls.TextBlock
        $line.Margin = '0,0,0,6'
        $line.Text = '{0}: {1}' -f $prop.Name, $prop.Value
        $detailsPanel.Children.Add($line) | Out-Null
    }
    Add-Tab 'Details' $detailsScroll

    $diagScroll = New-Object System.Windows.Controls.ScrollViewer
    $diagPanel = New-Object System.Windows.Controls.StackPanel
    $diagPanel.Margin = '12'
    $diagScroll.Content = $diagPanel
    foreach ($group in (Get-ItTasks | Group-Object Category | Sort-Object Name)) {
        $heading = New-Object System.Windows.Controls.TextBlock
        $heading.Text = $group.Name
        $heading.FontSize = 17
        $heading.FontWeight = 'SemiBold'
        $heading.Margin = '0,10,0,6'
        $diagPanel.Children.Add($heading) | Out-Null
        foreach ($task in $group.Group) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Margin = '0,2,0,4'
            $cb.IsChecked = $task.Recommended
            $cb.Content = '{0}{1}' -f $task.Name, ($(if ($task.AdminRequired) { '  [Admin required]' } else { '' }))
            if ($task.AdminRequired -and -not $Script:IsAdmin) {
                $cb.Foreground = '#b42318'
                $cb.ToolTip = 'This task will be skipped unless the tool is elevated.'
            }
            $Script:TaskCheckBoxes[$task.Id] = $cb
            $diagPanel.Children.Add($cb) | Out-Null
        }
    }
    Add-Tab 'Diagnostics' $diagScroll

    $reportsPanel = New-Object System.Windows.Controls.StackPanel
    $reportsPanel.Margin = '12'
    foreach ($text in @(
        "Run path: $Script:RunPath",
        "Report: $Script:ReportPath",
        "Reports folder copy: $(Join-Path $Script:ReportAssetPath 'full-report.html')",
        "Summary: $Script:SummaryPath",
        "Task log: $Script:TaskLogPath",
        "Transcript: $Script:TranscriptPath"
    )) {
        $line = New-Object System.Windows.Controls.TextBlock
        $line.Margin = '0,0,0,8'
        $line.Text = $text
        $reportsPanel.Children.Add($line) | Out-Null
    }
    Add-Tab 'Reports' $reportsPanel

    $settingsPanel = New-Object System.Windows.Controls.StackPanel
    $settingsPanel.Margin = '12'
    foreach ($text in @(
        "Output root: $RootPath",
        'Recommended runs exclude admin-only repair or reset actions.',
        'Installed applications are collected from registry uninstall keys, not Win32_Product.',
        'No email, SMTP, credential, or client-specific settings are stored.'
    )) {
        $line = New-Object System.Windows.Controls.TextBlock
        $line.Margin = '0,0,0,8'
        $line.Text = $text
        $settingsPanel.Children.Add($line) | Out-Null
    }
    Add-Tab 'Settings' $settingsPanel

    $aboutPanel = New-Object System.Windows.Controls.StackPanel
    $aboutPanel.Margin = '12'
    foreach ($text in @(
        "$AppName $Script:Version",
        'Lightweight Windows troubleshooting and reporting tool for support engineers.',
        'PowerShell 5.1+ WPF MVP.'
    )) {
        $line = New-Object System.Windows.Controls.TextBlock
        $line.Margin = '0,0,0,8'
        $line.Text = $text
        $aboutPanel.Children.Add($line) | Out-Null
    }
    Add-Tab 'About' $aboutPanel

    $Script:Ui.RunSelectedButton.Add_Click({ Invoke-SelectedTasks -RecommendedOnly:$false })
    $Script:Ui.RunRecommendedButton.Add_Click({ Invoke-SelectedTasks -RecommendedOnly:$true })
    $Script:Ui.OpenReportButton.Add_Click({ if (Test-Path -LiteralPath $Script:ReportPath) { Start-Process $Script:ReportPath } })
    $Script:Ui.OpenFolderButton.Add_Click({ Start-Process $Script:RunPath })
    $Script:Ui.ZipButton.Add_Click({
        $zipPath = "$Script:RunPath.zip"
        Compress-Archive -Path (Join-Path $Script:RunPath '*') -DestinationPath $zipPath -CompressionLevel Fastest -Force
        [System.Windows.MessageBox]::Show("Created ZIP:`n$zipPath", $AppName, 'OK', 'Information') | Out-Null
    })
    $window.Add_Closed({ if ($Script:TranscriptStarted) { Stop-Transcript | Out-Null } })
    return $window
}

try {
    $runInfo = @"
AppName: $AppName
Version: $Script:Version
Computer: $Script:ComputerName
User: $Script:CurrentUser
Admin: $Script:IsAdmin
RunPath: $Script:RunPath
Started: $($Script:RunStarted.ToString('o'))
"@
    Write-TextFile (Join-Path $Script:RunPath 'run-info.txt') $runInfo
    $window = New-Ui
    $window.ShowDialog() | Out-Null
} catch {
    Write-Error $_
    try { if ($Script:TranscriptStarted) { Stop-Transcript | Out-Null } } catch {}
    throw
}
