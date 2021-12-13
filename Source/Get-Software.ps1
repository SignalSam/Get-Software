<#
 .Synopsis
  Script that provides inventory of installed (MSI's via the registry) software on one or more systems.

 .Description
  Script utilizes current approved methods to gather an inventory of software which utilizes the Windows
  Installer portion of the registry. Can target one or more systems using multiple sets of credentials.
  Results can be captured as host output, simple text, CSV data, or Microsoft Excel document (done via COM
  and requires Microsoft Excel (2016 or later) to be installed on the system running the script). If the
  'Excel' export argument is presented and Microsoft Excel is not installed it will fail-over to a CSV export.

 .Parameter ExportType
  Screen: Directs output to the host running the script.
  Text: Writes output using the same format as the 'Screen' option to a .txt file.
  CSV:  Writes output in standard CSV format to a .csv file.
  Excel: Writes output in Excel 2016+ format to a .xlsx file.
  (All except 'Screen': If -OutFile was omitted, will attempt to create the file in the same directory as the script.)

 .Example
  PS C:\Get-Software -ExportType Text -RemoteHosts @('HOST01', 'HOST02', 'HOST03')

 .Example
  PS C:\Get-Software -ExportType CSV -Local -RemoteHosts @('HOST01', 'HOST02')

 .Example
  PS C:\Get-Software -ExportType Excel -Local -RemoteHosts $MyHostnames -CredentialSets 4

 .Link
  https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/aa394378(v=vs.85)
  https://docs.microsoft.com/en-US/troubleshoot/windows-server/admin-development/windows-installer-reconfigured-all-applications

 .Notes
  Configuration Managed Enterprises
  This script is not required for a Configuration Managed Enterprise, that tool has the functionality
  to provide this type of data. CM also creates and utilizes two new WMI classes that are query optimized.
  This script is intended for non-CM environments or for occasional use as required.

  Bad PowerShell Practices
  This was done also to highlight two important bad practices that are still pervasive in PowerShell use:
  1) Get-WmiObject is DEPRECIATED, it has been for quite some time. Use: Get-CimInstance
  2) Querying Win32_Product is a practice that Microsoft has outright said to avoid:
    Win32_Product Class
    https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/aa394378(v=vs.85)
    "Warning: Win32_Product is not query optimized. Queries such as
    "select * from Win32_Product where (name like 'Sniffer%')" require WMI to use the MSI provider to
    enumerate all of the installed products and then parse the full list sequentially to handle the
    “where” clause. This process also initiates a consistency check of packages installed, verifying and
    repairing the install. With an account with only user privileges, as the user account may not have
    access to quite a few locations, may cause delay in application launch and an event 11708 stating an
    installation failure. For more information, see KB Article 794524."

    (KB794524) Event log message indicates that the Windows Installer reconfigured all installed applications
    https://docs.microsoft.com/en-US/troubleshoot/windows-server/admin-development/windows-installer-reconfigured-all-applications
    Cause
    This problem can happen if one of the following conditions is true:
        You have a group policy with a WMIFilter that queries Win32_Product class.
        You have an application installed on the machine that queries Win32_Product class.

    Resolution
    If you're using a group policy with the WMIFilter that queries Win32_Product, modify the filter to use Win32reg_AddRemovePrograms.
    If you have an application that uses the previous class, contact the vendor to get an updated version that doesn't use this class.

    (Win32reg_AddRemovePrograms and its 64-bit equivalent only exist within a CM Enterprise.)

  Last updated: 12 December 2021
#>

[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $False)][Alias('E')][ValidateSet('Screen', 'Text', 'CSV', 'Excel')][String] $ExportType = 'Screen',
    [Parameter(Mandatory = $False)][Alias('L')][Switch] $Local,
    [Parameter(Mandatory = $False, ParameterSetName = 'HostsByData')][Alias('R')][System.Collections.Generic.List[String]] $RemoteHosts = @(),
    [Parameter(Mandatory = $False, ParameterSetName = 'HostsByFile')][Alias('I')][String] $InFile = '',
    [Parameter(Mandatory = $False)][Alias('O')][String] $OutFile = '',
    [Parameter(Mandatory = $False)][Alias('C')][Int32] $CredentialSets = 0
)

Function Test-RemoteHosts
{
    Param
        ( [Parameter(Mandatory = $True)][System.Collections.Generic.List[String]] $HostsToTest )

    [RegEx] $HostnameExpression = '^(?!\.)[A-Za-z0-9\.]{1,15}$'
    $RemoteHosts = $RemoteHosts | Select-Object -Unique

    ForEach ($HostFound In $HostsToTest)
    {
        If (-NOT ($HostnameExpression.IsMatch($HostFound)))
            { Return $HostFound }
    }

    Return $True
}

Function Test-OutfilePath
{

}

Function Test-Arguments
{
    If ($RemoteHosts.Count -GT 0)
    {
        $TestResult = Test-RemoteHosts $RemoteHosts

        If ($TestResult -NE $True)
        {
            Write-Error -Message "Invalid hostname detected [$TestResult] in -RemoteHosts input." -Category ([System.Management.Automation.ErrorCategory]::InvalidArgument)
            Exit
        }
    }
}

Test-Arguments