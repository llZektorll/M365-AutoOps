<#
.SYNOPSIS
    This script creates a self-signed certificate for use in M365 Auto Ops.

.DESCRIPTION
    Script developed to create and install a self-signed certificate to be used in Azure Application Authentication.

.NOTES
    Author      : Hugo Santos
    Date        : 2025-07-21
    Version     : 1.0.0
    Revision History:
        - 1.0.0 [2025-07-21] - Initial release
#>

#region Module Imports
Import-Module "C:\M365-AutoOps\M365 Auto Ops\M365AutoOps.psm1"
#endregion
#region Parameters

param (
    [Parameter(Mandatory = $True)]
    [string]$Message
)
#endregion
#region Global Variables
$script:Source = $MyInvocation.MyCommand.Path
#endregion

#region Variables
$Certificate_Name = 'M365 Auto Ops Certificate'
$Certificate_Description = 'M365 Auto Ops certificate for signing scripts.'
$Certificate_File = 'M365AutoOps_Certificate.cer'
$Certificate_Life = 5 # Certificate valid for X years

#Export
$Export = "$($Export_Path)\ToolKit\$($Certificate_File)"
#endregion

Try {
    #Terminal and Log write
    Write-LogEntry -Message "Creating the self-signed certificate" -Level 2  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
    Write-Host "Creating the self-signed certificate" -ForegroundColor Yellow
                
    #Execution
                $Certificate = New-SelfSignedCertificate -CertStoreLocation cert:\currentuser\my -Subject $Certificate_Name -KeyDescription $Certificate_Description -NotAfter (Get-Date).AddYears($Certificate_Life)
                $Certificate.Thumbprint | Clip
                Export-Certificate -Cert $Certificate -FilePath $Export -Type CERT
    #Terminal and Log write
    Write-LogEntry -Message "Self-signed certificate created" -Level 3  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
    Write-Host "Self-signed certificate created" -ForegroundColor Yellow
} Catch {
    Write-LogEntry -Message "Unable to create a self-signed certificate" -Level 5  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
}
$input = Read-Host "Do you want to check the scripts repository? (Press 'n' to cancel)"
if ($input -ne 'n') {
    Start-Process 'https://github.com/llZektorll/M365-AutoOps'
} else {
    break
}
