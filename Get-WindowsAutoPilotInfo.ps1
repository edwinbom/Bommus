<#PSScriptInfo

.VERSION 2.1
.GUID ebf446a3-3362-4774-83c0-b7299410b63f
.AUTHOR Michael Niehaus
.TAGS Windows AutoPilot
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$False,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=0)]
    [alias("DNSHostName","ComputerName","Computer")]
    [String[]] $Name = @("localhost"),

    [Parameter(Mandatory=$False)] [String] $OutputFile = "", 
    [Parameter(Mandatory=$False)] [String] $GroupTag = "",
    [Parameter(Mandatory=$False)] [Switch] $Append = $false,
    [Parameter(Mandatory=$False)] [System.Management.Automation.PSCredential] $Credential = $null,
    [Parameter(Mandatory=$False)] [Switch] $Partner = $false,
    [Parameter(Mandatory=$False)] [Switch] $Force = $false,
    [Parameter(Mandatory=$False)] [Switch] $Online = $false,
    [Parameter(Mandatory=$False)] [Switch] $Silent = $false
)

Begin {
    $computers = @()

    if ($Online) {
        $module = Import-Module WindowsAutopilotIntune -PassThru -ErrorAction Ignore
        if (-not $module) {
            if (-not $Silent) { Write-Host "Installing module WindowsAutopilotIntune" }
            Install-Module WindowsAutopilotIntune -Force
        }
        Import-Module WindowsAutopilotIntune -Scope Global
        $graph = Connect-MSGraph
        if (-not $Silent) { Write-Host "Connected to tenant $($graph.TenantId)" }

        if ($OutputFile -eq "") {
            $OutputFile = "$($env:TEMP)\autopilot.csv"
        } 
    }
}

Process {
    foreach ($comp in $Name) {
        $bad = $false

        if ($comp -eq "localhost") {
            $session = New-CimSession
        } else {
            $session = New-CimSession -ComputerName $comp -Credential $Credential
        }

        if (-not $Silent) { Write-Verbose "Checking $comp" }
        $serial = (Get-CimInstance -CimSession $session -Class Win32_BIOS).SerialNumber
        $devDetail = (Get-CimInstance -CimSession $session -Namespace root/cimv2/mdm/dmmap -Class MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'")

        if ($devDetail -and (-not $Force)) {
            $hash = $devDetail.DeviceHardwareData
        } else {
            $bad = $true
            $hash = ""
        }

        if ($bad -or $Force) {
            $cs = Get-CimInstance -CimSession $session -Class Win32_ComputerSystem
            $make = $cs.Manufacturer.Trim()
            $model = $cs.Model.Trim()
            if ($Partner) { $bad = $false }
        } else {
            $make = ""
            $model = ""
        }

        $product = ""

        if ($Partner) {
            $c = New-Object psobject -Property @{
                "Device Serial Number" = $serial
                "Windows Product ID"   = $product
                "Hardware Hash"        = $hash
                "Manufacturer name"    = $make
                "Device model"         = $model
            }
        }
        elseif ($GroupTag -ne "") {
            $c = New-Object psobject -Property @{
                "Device Serial Number" = $serial
                "Windows Product ID"   = $product
                "Hardware Hash"        = $hash
                "Group Tag"            = $GroupTag
            }
        }
        else {
            $c = New-Object psobject -Property @{
                "Device Serial Number" = $serial
                "Windows Product ID"   = $product
                "Hardware Hash"        = $hash
            }
        }

        if ($bad) {
            if (-not $Silent) {
                Write-Error -Message "Unable to retrieve device hardware data (hash) from computer $comp" -Category DeviceError
            }
        }
        elseif ($OutputFile -eq "") {
            $c
        }
        else {
            $computers += $c
        }

        Remove-CimSession $session
    }
}

End {
    if ($OutputFile -ne "") {
        if ($Append) {
            if (Test-Path $OutputFile) {
                $computers += Import-CSV -Path $OutputFile
            }
        }

        if ($Partner) {
            $computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash", "Manufacturer name", "Device model" |
                ConvertTo-CSV -NoTypeInformation | ForEach-Object { $_ -replace '"','' } | Out-File $OutputFile
        }
        elseif ($GroupTag -ne "") {
            $computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" |
                ConvertTo-CSV -NoTypeInformation | ForEach-Object { $_ -replace '"','' } | Out-File $OutputFile
        }
        else {
            $computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash" |
                ConvertTo-CSV -NoTypeInformation | ForEach-Object { $_ -replace '"','' } | Out-File $OutputFile
        }
    }

    if ($Online) {
        Import-AutopilotCSV -csvFile $OutputFile
    }
}
