<#
.SYNOPSIS
    This module contais all the functions needed to run M365 AutoOps as a single script.
.DESCRIPTION
    The module provides a set of functions to ensure that the M365 AutoOps script runs smoothly, including error handling, logging, and configuration management.
    It is designed to be used in conjunction with the M365 AutoOps script to automate various tasks in a Microsoft 365 environment.
.AUTHOR
    Author: Hugo Santos
    GitHub: https://github.com/llZektorll/M365-AutoOps
.VERSION
    0.1.0
.LASTUPDATED
    2025-07-19
.REQUIREMENTS
    - PowerShell 5.1+
    - Admin previleges on the machine
    - Microsoft 365 PowerShell modules installed
    - Internet connection for module updates
    - Access to Microsoft 365 tenant
    - Appropriate permissions in the Microsoft 365 tenant to run the script
.VARIABLES
    - All the variables used in the script are defined assuming that all the scripts and modules are store in C:\M365AutoOps
.NOTES 
    Creation Date: 2025-02-28
#>

#region variables
$Root_Path = "C:\M365AutoOps"
$Log_Path = "$($Root_Path)\Logs"
$Module_Path = "$($Root_Path)\Modules"
$Export_Path = "$($Root_Path)\Exports"

# Global Hashtable to store script execution IDs
# This allows tracking of script executions across different runs and sources
if (-not $global:ScriptExecutionIDs) {
    $global:ScriptExecutionIDs = @{}
}
#endregion

#region Functions
#Logging function
Function Write-LogEntry {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]$Message,
        
        [Parameter()]
        [string]$Error_Command,

        [Parameter()]
        [string]$Error_Line,
        
        [Parameter()]
        [string]$Error_Message,

        [Parameter(Mandatory = $True)]
        [ValidateSet(1,2,3,4,5)]
        [string]$Level,

        [Parameter()]
        [ValidateSet(1,2,3,4,5)]
        [string]$MinimumLogLevel = 1,

        [Parameter()]
        [string]$Log = "$($Log_Path)\Logging_$(Get-Date -Format 'yyyyMMdd').csv",

        [Parameter()]
        [string]$Source
    )

    # Write header if file does not exist
    if (-not (Test-Path $Log_Path)) {
        $Header = "Timestamp;Level;Source;ExecutionID;Message;Error_Command;Error_Line;Error_Message"
        $Header | Out-File -FilePath $Log_Path -Encoding UTF8
    }

    #Ensure the log level is valid
    $Log_Level = @{
        'DEBUG'   = 1
        'INFO'    = 2
        'SUCCESS' = 3
        'WARNING' = 4
        'ERROR'   = 5
    }
    foreach ($Log_LVL in $Log_Level.GetEnumerator()) {
        if ($Log_LVL.Value -eq $Level) {
            $Save_Log_Level = $Log_LVL.Key
            break
        }
    }
    If (-not($Save_Log_Level -match "ERROR")) {
        $Error_Command = ""
        $Error_Line = ""
        $Error_Message = ""
    }
    #Ensure the sorurce is valid
    If (-not $Source) {
        $Source = "Null"
    }
    #Ensure the Execution ID is set for the source
    if (-not $global:ScriptExecutionIDs.ContainsKey($Source)) {
        $global:ScriptExecutionIDs[$Source] = [guid]::NewGuid().ToString()
    }
    $ExecutionID = $global:ScriptExecutionIDs[$Source]
    
    #Export the log entry to a CSV file
    $Log_Entry = [PSCustomObject][Ordered]@{
        Timestamp     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Level         = $Save_Log_Level
        ExecutionID   = $ExecutionID
        Source        = $Source
        Message       = $Message
        Error_Command = $Error_Command
        Error_Line    = $Error_Line
        Error_Message = $Error_Message
    }
    $Log_Entry | Export-Csv -Path $Log -Append -Delimiter ";" -Encoding utf8 -NoClobber -Force:$True -Confirm:$false 

    #Logging example
    #Import-Module "C:\M365AutoOps\M365 Auto Ops\M365AutoOps.psm1"
    #$Source = $MyInvocation.MyCommand.Path
    #Write-LogEntry -Message "teste5" -Level 5  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
}
#Enseure the paths exist
Function Path-Finder {
    If (-not (Test-Path -Path $Root_Path)) {
        New-Item -Path $Root_Path -ItemType Directory -Force
    }
    If (-not (Test-Path -Path $Log_Path)) {
        New-Item -Path $Log_Path -ItemType Directory -Force
    }
    If (-not (Test-Path -Path $Module_Path)) {
        New-Item -ItemType Directory -Path $Module_Path -Force
    }
    If (-not (Test-Path -Path $Export_Path)) {
        New-Item -ItemType Directory -Path $Export_Path -Force
    }
}
#Ensure TLS 1.2 is enabled
Function Ensure-TLS12 {
    if ([Net.ServicePointManager]::SecurityProtocol -ne [Net.SecurityProtocolType]::Tls12) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }
}

#endregion

#region default execution
Path-Finder
Ensure-TLS12
#endregion