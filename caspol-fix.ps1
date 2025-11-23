# Enable .NET 4 Legacy CAS Policy and reset CASPOL
# Works for both 32-bit and 64-bit .NET Framework
# Run this script in an elevated PowerShell window (Run as Administrator)

$ErrorActionPreference = "Stop"

# Paths
$windowsDir   = [System.Environment]::ExpandEnvironmentVariables("%WINDIR%")
$framework64  = Join-Path $windowsDir "Microsoft.NET\Framework64"
$framework32  = Join-Path $windowsDir "Microsoft.NET\Framework"

$net40_64     = Join-Path $framework64 "v4.0.30319\Config\machine.config"
$net40_32     = Join-Path $framework32 "v4.0.30319\Config\machine.config"

function Enable-LegacyCAS {
    param (
        [string] $MachineConfigPath,
        [string] $CaspolPath
    )

    if (-Not (Test-Path $MachineConfigPath)) {
        Write-Host "File not found: $MachineConfigPath"
        return
    }

    # Backup
    Write-Host "Backing up $MachineConfigPath ..."
    Copy-Item $MachineConfigPath "$MachineConfigPath.bak" -Force

    # Load XML
    [xml]$xml = Get-Content $MachineConfigPath

    # Ensure <runtime> exists
    $runtimeNode = $xml.configuration.runtime

    if (-not $runtimeNode) {
        $runtimeNode = $xml.CreateElement("runtime")
        $xml.configuration.AppendChild($runtimeNode) | Out-Null
    }

    # Check if NetFx40_LegacySecurityPolicy exists
    $legacyNode = $xml.configuration.runtime.NetFx40_LegacySecurityPolicy
    if (-not $legacyNode) {
        $legacyNode = $xml.CreateElement("NetFx40_LegacySecurityPolicy")
        $legacyNode.SetAttribute("enabled", "true")
        $runtimeNode.AppendChild($legacyNode) | Out-Null
        $xml.Save($MachineConfigPath)
        Write-Host "Legacy CAS policy enabled in machine.config"
    } else {
        $legacyNode.SetAttribute("enabled", "true")
        $xml.Save($MachineConfigPath)
        Write-Host "Legacy CAS policy already present"
    }

    #Reset CAS policy
    if (Test-Path $CaspolPath) {
        Write-Host "Resetting CAS policy with $CaspolPath ..."
        & $CaspolPath -polchgprompt off
        & $CaspolPath -all -reset
        & $CaspolPath -polchgprompt on
        Write-Host "CAS policy reset completed."
    } else {
        Write-Host "CASPOL not found at $CaspolPath"
    }

}

# Enable for 64-bit .NET if present
if (Test-Path $net40_64) {
    $caspol64 = Join-Path $framework64 "v4.0.30319\caspol.exe"
    Enable-LegacyCAS -MachineConfigPath $net40_64 -CaspolPath $caspol64
}

# Enable for 32-bit .NET if present
if (Test-Path $net40_32) {
    $caspol32 = Join-Path $framework32 "v4.0.30319\caspol.exe"
    Enable-LegacyCAS -MachineConfigPath $net40_32 -CaspolPath $caspol32
}

Write-Host "PROCESS COMPLETE!"
