<#
.SYNOPSIS
    Microsoft 365 AutoOps - Microsoft 365 All PowerShell Modules Installer

.DESCRIPTION
    Script developed to automatically install or download all PowerShell modules for Microsoft 365

.NOTES
    Author      : Hugo Santos
    Date        : 2025-07-21
    Version     : 1.0.0
    Revision History:
        - 1.0.0 [2025-07-21] - Initial release
#>

#region Module Imports
Import-Module "C:\M365AutoOps\M365 Auto Ops\M365AutoOps.psm1"
#endregion

#region Global Variables
$script:Source = $MyInvocation.MyCommand.Path
#endregion

#region Variables
$M365_Modules = @(
    'ExchangeOnlineManagement',
    'Microsoft.Graph',
    'MicrosoftTeams',
    'Microsoft.Online.SharePoint.PowerShell',
    'PNP.PowerShell',
    'Microsoft.PowerApps.Administration.PowerShell',
    'MicrosoftPowerBIMgmt'
)
#endregion

Write-Host "----------------------------------------------" -ForegroundColor Cyan
Write-Host "---                                        ---" -ForegroundColor Cyan
Write-Host "---           M365 Auto Ops                ---" -ForegroundColor Cyan
Write-Host "---                                        ---" -ForegroundColor Cyan
Write-Host "----------------------------------------------" -ForegroundColor Cyan
Write-Host "---      1 - Install All Modules           ---" -ForegroundColor Cyan
Write-Host "---      2 - Download All Modules          ---" -ForegroundColor Cyan
Write-Host "----------------------------------------------" -ForegroundColor Cyan
$Execution_Type = Read-Host "Do you want to install or download all the modules? [1/2]"

If ($Execution_Type -like 1) {
    Try {
        Write-LogEntry -Message "Installing all Modules " -Level 2  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
        Foreach ($Module in $M365_Modules) {
            Try {
                #Terminal and Log write
                Write-LogEntry -Message "Installing $($Module)" -Level 2  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
                Write-Host "Installing $($Module)" -ForegroundColor Yellow
                
                #Module installation
                Install-Module -Name $Module -Confirm:$false -Force
                
                #Terminal and Log write
                Write-LogEntry -Message "Module $($Module) installed successfully" -Level 3  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
                Write-Host "Module $($Module) installed successfully" -ForegroundColor Yellow
            } Catch {
                Write-LogEntry -Message "Unable to install $($Module)" -Level 5  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
                Write-Host "Unable to install $($Module)" -ForegroundColor Red
            }
        }
        Write-LogEntry -Message "Installation Completed " -Level 3  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
    } Catch {
        Write-LogEntry -Message "Unable to run the installation process" -Level 5  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
    }
}ElseIf($Execution_Type -like 2){
    Try {
        Write-LogEntry -Message "Downloading all Modules " -Level 2  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
        Foreach ($Module in $M365_Modules) {
            Try {
                #Terminal and Log write
                Write-LogEntry -Message "Downloading $($Module)" -Level 2  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
                Write-Host "Downloading $($Module)" -ForegroundColor Yellow
                
                #Module installation
                Install-Module -Name $Module -Confirm:$false -Force
                
                #Terminal and Log write
                Write-LogEntry -Message "Module $($Module) Downloaded successfully" -Level 3  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
                Write-Host "Module $($Module) Downloaded successfully" -ForegroundColor Yellow
            } Catch {
                Write-LogEntry -Message "Unable to Download $($Module)" -Level 5  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
                Write-Host "Unable to Download $($Module)" -ForegroundColor Red
            }
        }
        Write-LogEntry -Message "Download Completed " -Level 3  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
    } Catch {
        Write-LogEntry -Message "Unable to run the Download process" -Level 5  -Source $Source -Error_Command "$($_.InvocationInfo.InvocationName)" -Error_Line "$($_.InvocationInfo.ScriptLineNumber)" -Error_Message "$($_.Exception.Message)"
    }
}Else{
    Write-Host "Invalid Option" -ForegroundColor Red
    Write-Host "Closing PowerShell" -ForegroundColor Red
}
$input = Read-Host "Do you want to check the scripts repository? (Press 'n' to cancel)"
if ($input -ne 'n') {
    Start-Process 'https://github.com/llZektorll/M365-AutoOps'
} else {
    break
}
