# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NonInteractive -File D:\Scripts\Report-PrintServerIssues\Report-PrintServerIssues.ps1

[CmdletBinding(DefaultParameterSetName="Schedule")]
Param (

    [Parameter(Mandatory=$False, ParameterSetName='Schedule')]
    [int]$SendEmailDay = 1,

    [Parameter(Mandatory=$False, ParameterSetName='EmailForce')]
    [switch]$ForceEmail

)


Function Append-HtmlSection {

    Param (

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        $Table,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$Heading,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [string]$Reason,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [string]$Action

    )

    $HtmlArgs = @{}
    $HtmlArgs['PreContent'] = "<h2>$Heading</h2>$(If ($Reason) {"<p>$Reason."})<p>$($Table.Count) issue(s) found:<p>"
    If ($Action) {
        $HtmlArgs['PostContent'] = "<p><i>Corrective action: $($Action).</i>"
    }
    ($Table | ConvertTo-Html -Fragment @HtmlArgs) -replace '###', '<br>' | Out-File $HtmlReportPathLocal -Encoding default -Append

}



# --------------------------------------------------
# Define variables
# --------------------------------------------------

$EmailSubject = "Weekly report on print server issues"
$EmailSender = "PrintServerIssues@domain.tld"
$EmailServer = "smtp.domain.tld"

$ScriptBasePath = Get-Item $PSCommandPath
$ReportBasePath = "D:\Reports\PrintServerIssues"

If (-Not(Test-Path $ScriptBasePath -ErrorAction SilentlyContinue)) {
    Throw "Base path ""$ScriptBasePath"" inaccesible"
}

$PrintServersTxt = "$ScriptBasePath\PrintServers.txt"
$EmailRecipientsTxt = "$ScriptBasePath\EmailRecipients.txt"
$CommonDriversTxt = "$ScriptBasePath\CommonDrivers.txt"
$IgnoreDriversTxt = "$ScriptBasePath\IgnoreDrivers.txt"
$IgnorePrintersTxt = "$ScriptBasePath\IgnorePrinters.txt"
$UniDriversCsv = "$ScriptBasePath\UniDrivers.csv"
$HtmlReportPathLocal = "$ReportBasePath\PrintServerIssues.html"
$HtmlReportPathRemote = $HtmlReportPathLocal.Replace((Get-Item $HtmlReportPathLocal).PSDrive.Root,"\\$($env:COMPUTERNAME.ToLower())\")
$RegKey = 'HKCU:\SOFTWARE\DFDS\Report-PrintServerIssues' # Do not run PS with -NoProfile



# --------------------------------------------------
# Load server and email recipient lists
# --------------------------------------------------

$ComputerName = Get-Content $PrintServersTxt
$EmailRecipient = Get-Content $EmailRecipientsTxt



# --------------------------------------------------
# Load lists of common drivers and stuff to ignore
# --------------------------------------------------

$CommonDrivers = Get-Content $CommonDriversTxt
$IgnorePrinters = Get-Content $IgnorePrintersTxt
$IgnoreDrivers = Get-Content $IgnoreDriversTxt
$ProcessorArchitectures = @('Windows NT x86', 'Windows x64')



# --------------------------------------------------
# Load uni driver mapping
# --------------------------------------------------

$UniDrivers = @{}
Import-Csv $UniDriversCsv -Delimiter ';' | ForEach {
    $UniDrivers[$_.Brand] = $_.DriverName
}



# --------------------------------------------------
# Test if each server is reachable
# --------------------------------------------------

# Define ping status messages
$PingStatusTable = @{}
$PingStatusTable['0'] = 'Success'
$PingStatusTable['11001'] = 'Buffer Too Small'
$PingStatusTable['11002'] = 'Destination Net Unreachable'
$PingStatusTable['11003'] = 'Destination Host Unreachable'
$PingStatusTable['11004'] = 'Destination Protocol Unreachable'
$PingStatusTable['11005'] = 'Destination Port Unreachable'
$PingStatusTable['11006'] = 'No Resources'
$PingStatusTable['11007'] = 'Bad Option'
$PingStatusTable['11008'] = 'Hardware Error'
$PingStatusTable['11009'] = 'Packet Too Big'
$PingStatusTable['11010'] = 'Request Timed Out'
$PingStatusTable['11011'] = 'Bad Request'
$PingStatusTable['11012'] = 'Bad Route'
$PingStatusTable['11013'] = 'TimeToLive Expired Transit'
$PingStatusTable['11014'] = 'TimeToLive Expired Reassembly'
$PingStatusTable['11015'] = 'Parameter Problem'
$PingStatusTable['11016'] = 'Source Quench'
$PingStatusTable['11017'] = 'Option Too Big'
$PingStatusTable['11018'] = 'Bad Destination'
$PingStatusTable['11032'] = 'Negotiating IPSEC'
$PingStatusTable['11050'] = 'General Failure'

# Test connection
$ComputerError = @()
ForEach ($Computer in ($ComputerName | Sort-Object)) {

    # Attempt to lookup computername in DNS, add to $ComputerError array with error message if fails
    Try {$IPAddress = @(Resolve-DnsName -Name $Computer -ErrorAction Stop)[0] | Select -Expand IPAddress; $Resolvable = $true}
    Catch {
        $ErrorMessage = "$($_.InvocationInfo.MyCommand): $($_.Exception.Message)"
        $ComputerError += [pscustomobject]@{
            ComputerName = $Computer
            Error = $ErrorMessage
        }
        Write-Warning "$Computer - $ErrorMessage"
        Continue
    }


    # Attempt to ping computername, add to $ComputerError array with error message if fails
    $PingStatus = (Get-WmiObject -Class Win32_PingStatus -Filter "Address='$IPAddress' AND Timeout=1000").StatusCode
    If ($PingStatus -ne 0) {
        $ErrorMessage = "Win32_PingStatus: '$($PingStatusTable["$PingStatus"])' ($PingStatus)"
        $ComputerError += [pscustomobject]@{
            ComputerName = $Computer
            Error = $ErrorMessage
        }
        Write-Warning "$Computer - $ErrorMessage"
        Continue
    }


    # Check for remote admin rights
    If (-Not(Test-Path "\\dkcph-infpprt1\admin$" -ErrorAction SilentlyContinue)) {
        $ErrorMessage = "No admin rights on computer"
        $ComputerError += [pscustomobject]@{
            ComputerName = $Computer
            Error = $ErrorMessage
        }
        Write-Warning "$Computer - $ErrorMessage"
        Continue
    }

}



# --------------------------------------------------
# Get printers
# --------------------------------------------------

$ComputerNoError = $ComputerName | Where-Object {$ComputerError.ComputerName -notcontains $_} | Sort-Object
Write-Verbose "Getting printer information from $($ComputerNoError.Count) computer(s)" -Verbose

$Printers = ForEach ($Computer in $ComputerNoError) {

    Try {Get-Printer -ComputerName $Computer -ErrorAction Stop | Where-Object {$IgnorePrinters -notcontains $_.Name} | Sort-Object Name}

    Catch {
        $ErrorMessage = "$($_.InvocationInfo.MyCommand): $($_.Exception.Message)"
        $ComputerError += [pscustomobject]@{
            ComputerName = $Computer
            Error = $ErrorMessage
        }
        Write-Warning "$Computer - $ErrorMessage"
    }

}



# --------------------------------------------------
# Get drivers
# --------------------------------------------------

$ComputerNoError = $ComputerName | Where-Object {$ComputerError.ComputerName -notcontains $_} | Sort-Object
Write-Verbose "Getting driver information from $($ComputerNoError.Count) computer(s)" -Verbose

$Drivers = ForEach ($Computer in $ComputerNoError) {

    Try {Get-PrinterDriver -ComputerName $Computer | Where-Object {$IgnoreDrivers -notcontains $_.Name} | Sort-Object ComputerName, Name, PrinterEnvironment}

    Catch {
        $ErrorMessage = "$($_.InvocationInfo.MyCommand): $($_.Exception.Message)"
        $ComputerError += [pscustomobject]@{
            ComputerName = $Computer
            Error = $ErrorMessage
        }
        Write-Warning "$Computer - $ErrorMessage"
    }

}



# --------------------------------------------------
# Get ports
# --------------------------------------------------

$Ports = @()
$ComputerNoError = $ComputerName | Where-Object {$ComputerError.ComputerName -notcontains $_} | Sort-Object
Write-Verbose "Getting port information from $($ComputerNoError.Count) computer(s)" -Verbose

$Ports = ForEach ($Computer in $ComputerNoError) {

    Try {Get-PrinterPort -ComputerName $Computer | Where-Object {$_.Description -ne 'Local Port'} | Sort-Object Name}

    Catch {
        $ErrorMessage = "$($_.InvocationInfo.MyCommand): $($_.Exception.Message)"
        $ComputerError += [pscustomobject]@{
            ComputerName = $Computer
            Error = $ErrorMessage
        }
        Write-Warning "$Computer - $ErrorMessage"
    }

}



# --------------------------------------------------
# Initiate report file
# --------------------------------------------------

$HtmlPre = @"
<title>Print server issues</title>
<meta http-equiv='refresh' content='900'>
<style>
    TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
    TH {border-width: 1px; padding: 5px; border-style: solid; border-color: black; background-color: #1B5786; text-align: left; color: white; font-family: Verdana; font-size: 0.875em}
    TD {border-width: 1px; padding: 5px; border-style: solid; border-color: black; background-color: #E3EAF0; font-family: Verdana; font-size: 0.875em}
    Body {font-family: Verdana; font-size: 0.875em}
    H1 {font-family: Verdana}
</style>
<body>
<h1>Print server issues</h1>

<h2>Summary</h2>
<p>Report generated on <b>$((Get-Date -Format s).Replace('T',' ').Substring(0,16))</b>.
<br>Always up-to-date report: <a href="file://$HtmlReportPathRemote">$HtmlReportPathRemote</a>.

<p>Queried <b>$($ComputerName.Count)</b> print servers: $(($ComputerName | Sort-Object) -join ', ')

<p>Found <b>$($Printers.Count)</b> printers
<br>Found <b>$($Drivers.Count)</b> drivers
<br>Found <b>$($Ports.Count)</b> ports

<p>Ignored <b>$($IgnorePrinters.Count)</b> printer(s): $(($IgnorePrinters | Sort-Object) -join ', ')
<br>Ignored <b>$($IgnoreDrivers.Count)</b> driver(s):  $(($IgnoreDrivers | Sort-Object) -join ', ')
</ul>
"@
$HtmlPre | Out-File $HtmlReportPathLocal -Encoding default



# --------------------------------------------------
# Computers with issues
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Computers with issues"
If ($ComputerError) {Append-HtmlSection -Heading $Heading -Table ($ComputerError | Sort-Object ComputerName)}



# --------------------------------------------------
# Computers with common drivers missing
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Computers with common drivers missing"
$Reason = "Common drivers should be installed on all print servers, to ensure new printers can easily be setup"
$Action = "Run script to install common drivers: \\dkcph-infpswl1\DML\Hardware\Print\Install-Common-Drivers.cmd"
Write-Verbose $Heading -Verbose
$Table = @(ForEach ($Computer in $ComputerNoError) {

    ForEach ($Driver in $CommonDrivers) { 

        $DriverInstalled = [bool]($Drivers | Where-Object {$_.ComputerName -eq $Computer -and $_.Name -eq $Driver})

        If (-Not($DriverInstalled)) {

            [pscustomobject]@{
                ComputerName = $Computer
                DriverName = $Driver
            }
            
        }

    }

})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Drivers not in use (except default ones)
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Drivers not in use"
$Reason = "Drivers not in use muddies the waters, increasing the risk of choosing a wrong driver for printers"
$Action = "Remove unnecessary drivers"
Write-Verbose $Heading -Verbose
$Table = @(ForEach ($Computer in ($Drivers | Where-Object {$CommonDrivers -notcontains $_.Name} | Select -Unique -ExpandProperty ComputerName)) {

    $ComputerDriverNames = $Drivers | Where-Object {$_.ComputerName -eq $Computer -and $CommonDrivers -notcontains $_.Name} | Select-Object -Unique -ExpandProperty Name

    ForEach ($Driver in $ComputerDriverNames) {

        $DriverInUse = [bool]($Printers | Where-Object {$_.ComputerName -eq $Computer -and $_.DriverName -eq $Driver})

        If (-Not($DriverInUse)) {

            [pscustomobject]@{
                ComputerName = $Computer
                DriverName = $Driver
            }

        }
        
    }

})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Drivers not installed for all processor architectures
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Drivers not installed for all processor architectures"
$Reason = "Drivers need to be installed for all supported architectures, to ensure that both 32-bit and 64-bit clients can print"
$Action = "Install drivers for the missing processor architectures"
Write-Verbose $Heading -Verbose
$DriversMissingArchitecture = @()
$Table = @(ForEach ($Computer in ($Drivers | Select -Unique -ExpandProperty ComputerName)) {

    $Drivers | Where-Object {$_.ComputerName -eq $Computer} | Group-Object Name | ForEach {

        $DriverName = $_.Name

        Compare-Object -ReferenceObject $ProcessorArchitectures -DifferenceObject $_.Group.PrinterEnvironment -PassThru | ForEach {

            $DriversMissingArchitecture += [pscustomobject]@{
                ComputerName = $Computer
                DriverName = $DriverName
                MissingArchitecture = $_
            }

        }

    }

})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Comments not populated
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Printers without comments"
$Reason = "The comment field must contain the make and model of the printer. This is also required for other print server health checks"
$Action = "Add/correct printer model in comment, or delete and re-create printer using Excel sheet"
Write-Verbose $Heading -Verbose
$Table = @($Printers | Where-Object {$_.Comment.Length -eq 0} | Select ComputerName, Name, DriverName)
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Printers not using proper drivers
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Printers using wrong drivers"
$Reason = "Wrong printer drivers can lead to printing issues"
$Action = "Add/correct printer model in comment"
Write-Verbose $Heading -Verbose
$Table = @(ForEach ($Printer in ($Printers | Where-Object {$_.Comment.Length -gt 0})) {

    $Brand = $Printer.Comment.Split(' ')[0]
    $UniDriver = $UniDrivers[$Brand]
    
    Switch ([bool]$UniDriver) {

        $true {
            $ExpectedDriverPrefix = $UniDriver
        }
        Default {
            $ExpectedDriverPrefix = $Brand
        }
    }

    If ($Printer.DriverName -notlike "$($ExpectedDriverPrefix)*") {

        [pscustomobject]@{
            ComputerName = $Printer.ComputerName
            PrinterName = $Printer.Name
            PrinterComment = $Printer.Comment
            DriverName = $Printer.DriverName
            ExpectedDriverPrefix = $ExpectedDriverPrefix
        }

    }

})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Ports not using proper names
# (indicating manual creation)
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Printer ports not using proper names"
$Reason = "Printer ports must be created using the Excel sheet to ensure consistency. The printer port name must correspond with the name of the printer"
$Action = "Re-create ports using Excel sheet, attach printers to correct ports"
Write-Verbose $Heading -Verbose
$Table = @($Ports | Where-Object {$_.Name -notlike '*-*'} | Select ComputerName, @{N='PortName';E={$_.Name}})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Port/print mismatch
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Printer/port name mismatch"
$Reason = "Printer ports must be created using the Excel sheet to ensure consistency. The printer port name must correspond with the name of the printer"
$Action = "Re-create ports using Excel sheet, attach printers to correct ports"
Write-Verbose $Heading -Verbose
$NonFollowMePrinters = $Printers | Where-Object {$_.Name -notlike '*Follow Me'}
$Table = @(ForEach ($Printer in $NonFollowMePrinters) {

    # Ignoring spaces and hyphens, does the printer name properly correlate with the port name?
    If (-Not("$($Printer.Name -replace "[ -&]", '')" -like "$($Printer.PortName -replace "[ -]", '')*")) {

        [pscustomobject]@{
            ComputerName = $Printer.ComputerName
            PrinterName = $Printer.Name
            PortName = $Printer.PortName
        }
        
    }

})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Ports not in use
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Ports not in use"
$Reason = "Ports not in use are probably left-overs after making changes. Remove them to avoid confusion"
$Action = "Remove ports not in use"
Write-Verbose $Heading -Verbose
$Table = @(ForEach ($Computer in ($Ports | Select -Unique -ExpandProperty ComputerName)) {

    $ComputerPortNames = $Ports | Where-Object {$_.ComputerName -eq $Computer} | Select-Object -Unique -ExpandProperty Name

    ForEach ($Port in $ComputerPortNames) {

        $PortInUse = [bool]($Printers | Where-Object {$_.ComputerName -eq $Computer -and $_.PortName -eq $Port})

        If (-Not($PortInUse)) {

            [pscustomobject]@{
                ComputerName = $Computer
                PortName = $Port
            }

        }
        
    }

})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Ports with more than one printer?
# Shortest name, match like* against others?
# --------------------------------------------------

Remove-Variable Heading, Reason, Action, Table -ErrorAction SilentlyContinue
$Heading = "Ports with more than one printer attached"
$Reason = "In some cases this could be due to legacy multi-port print servers being used. In most cases this is just left-overs after testing, or a simple mistake"
$Action = "Ensure printer-to-port mappings are correct"
Write-Verbose $Heading -Verbose
$Table = @(ForEach ($Computer in ($Ports | Select -Unique -ExpandProperty ComputerName)) {

    $ComputerPortNames = $Ports | Where-Object {$_.ComputerName -eq $Computer} | Select-Object -Unique -ExpandProperty Name

    ForEach ($Port in $ComputerPortNames) {

        $PortPrinters = @($Printers | Where-Object {$_.ComputerName -eq $Computer -and $_.PortName -eq $Port} | Select-Object -ExpandProperty Name)
        $ShortestName = $PortPrinters | Sort-Object Length | Select-Object -First 1


        If ($PortPrinters | Where-Object {$_ -notlike "$ShortestName*"}) {

            [pscustomobject]@{
                ComputerName = $Computer
                PortName = $Port
                PrinterNames = $PortPrinters -join "###"
            }

        }
        
    }

})
If ($Table) {Append-HtmlSection -Heading $Heading -Table $Table -Reason $Reason -Action $Action}



# --------------------------------------------------
# Wrap up HTML report
# --------------------------------------------------

$HtmlPost = @"
</body>
</html>
"@
$HtmlPost | Out-File $HtmlReportPathLocal -Encoding default -Append
Write-Verbose "Report generated" -Verbose



# --------------------------------------------------
# Send e-mail?
# --------------------------------------------------

# Debug
<#
Remove-ItemProperty $RegKey -Name 'EmailLastSent'
(Get-ItemProperty $RegKey -Name 'EmailLastSent').EmailLastSent
#>


# Any email recipients defined?
If ($EmailRecipient) {

    # Get time email was last sent, default to "very long ago"
    Try {$EmailLastSent = Get-Date (Get-ItemProperty $RegKey -Name 'EmailLastSent' -ErrorAction SilentlyContinue).EmailLastSent}
    Catch {$EmailLastSent = Get-Date 0}

    # Calculate days since email was last sent
    $Now = Get-Date
    $DayOfWeek = $Now.DayOfWeek.value__
    $HoursSinceLastEmail = New-TimeSpan -Start $EmailLastSent -End $Now | Select-Object -Expand TotalHours

    # If day of week matches $SendEmailDay and it's been at least 24 hours since last e-mail (to avoid spamming, if report is generated more frequent than daily), start send e-mail routine
    [bool]$SendEmail = ($ForceEmail -or ($DayOfWeek -eq $SendEmailDay -and $HoursSinceLastEmail -ge 24)) -and $EmailRecipient
    If ($SendEmail) {

        # Send report
        $EmailBody = Get-Content $HtmlReportPathLocal -Raw
        Send-MailMessage -Body $EmailBody -BodyAsHtml -Subject $EmailSubject -From $EmailSender -To $EmailRecipient -SmtpServer $EmailServer

        # Record time email was sent to registry
        If (-Not(Test-Path $RegKey)) {
            New-Item -Path $RegKey -Force | Out-Null
        }
        Set-ItemProperty $RegKey -Name 'EmailLastSent' -Value (Get-Date -Format 's') -Force
        Write-Verbose "Emailed report to $(($EmailRecipient | Sort-Object) -join ', ')" -Verbose

    } Else {

        Write-Verbose "Not emailing report. It is not $([DayOfWeek]$SendEmailDay) or report has already been emailed in the last 24 hours, and -ForceEmail was not specified." -Verbose

    }

}