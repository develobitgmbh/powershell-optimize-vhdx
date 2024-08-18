<#PSScriptInfo
.VERSION 1.0

.AUTHOR Marcel Jauernig

.COMPANYNAME develobit GmbH

.COPYRIGHT (c) 2024 Marcel Jauernig

.LICENSEURI https://github.com/develobitgmbh/powershell-optimize-vhdx/LICENSE.txt

.PROJECTURI https://github.com/develobitgmbh/powershell-optimize-vhdx

.RELEASENOTES
Version 1.0: Original published version.

#>

<#
.SYNOPSIS
This script tries to shrink vhdx files. Execute the script with administrator rights.

MIT License

Copyright (c) 2024 develobit GmbH

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

.DESCRIPTION
Execute the script with administrator rights.
This script tries to shrink vhdx files. It defragments the file and then tries to shrink the file by the unused disk space via diskpart.
The script has three phases. 
The first phase creates and runs a diskpart script to attach the vdisk. 
The second phase attach each partition and defragments the volume using the Powershell Optimize-Volume commandlet. 
The third phase creates a diskpart script to detach and shrink the disk. 
The script works without the Powershell commandlet Optimize-VHD and don't need the Hyper-V Management Tools.
The script is inspired by Johan Arwidmark https://www.deploymentresearch.com/optimizing-vhdx-files-in-a-hyper-v-lab/

.PARAMETER DiskPartScriptfile
The file name for the temporary diskpart script. Default file name is 'diskpart.txt'.

.PARAMETER Filter
The file filter to be used. *.vhdx for all vhdx files in the directory. sample.vhdx for a specific vhdx file. Default filter is '*.vhdx'.

.PARAMETER Letter
Drive letter that is used temporarily to mount the vdisk. Default letter is 'T'.

.PARAMETER Path
The path where the vhdx files are located.

.PARAMETER PauseTimeAfterEachStep
The time (in seconds) to be paused after each processing step. Default time are 3 seconds

.EXAMPLE
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles

.EXAMPLE
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -DiskPartScriptfile script.txt

.EXAMPLE
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -Letter J

.EXAMPLE
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -PauseTimeAfterEachStep 15

.EXAMPLE
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -Filter UVHD-S-1-5-21-2967749299-2196513631-28783610-1237.vhdx

.EXAMPLE
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -Filter UVHD-S-1-5-21-*.vhdx

.EXAMPLE
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -DiskPartScriptfile script.txt -Filter UVHD-S-1-5-21-2967749299-2196513631-28783610-1237.vhdx -Letter J -PauseTimeAfterEachStep 15

#>


Param(
   [Parameter(
    HelpMessage="The file name for the temporary diskpart script.",
    Mandatory=$false
   )] [string]$DiskPartScriptfile = "diskpart.txt",    
   [Parameter(
    HelpMessage="The file filter to be used. *.vhdx for all vhdx files in the directory. sample.vhdx for a specific vhdx file.",
    Mandatory=$false
   )] [string]$Filter = "*.vhdx",    
   [Parameter(
    HelpMessage="Drive letter that is used temporarily to mount the vdisk.",
    Mandatory=$false
   )] [string]$Letter = "T",    
   [Parameter(
    HelpMessage="The path where the vhdx files are located.",
    Mandatory=$true
   )] [string]$Path,
   [Parameter(
    HelpMessage="The time (in seconds) to be paused after each processing step.",
    Mandatory=$false
   )] [int]$PauseTimeAfterEachStep = 3
)

if (-not ($DiskPartScriptfile -match '^[A-Za-z0-9]+\.txt$')) {
    Write-Host "Error: The file name for the temporary diskpart script should only consist of letters and numbers. It must also end with '.txt'."
    Exit 1
}

if (Test-Path -Path $DiskPartScriptfile) {
    Write-Host "Error: The file '$DiskPartScriptfile' already exists."
    Exit 1
} 

if(-not ($Filter.ToLower().EndsWith(".vhdx"))) {
    Write-Host "Error: The filter does not end on '.vhdx'."
    Exit 1
}

if (-not ($Letter -match '^[A-Za-z]$')) {
    Write-Host "Error: The drive letter has not been specified as a single letter."
    Exit 1
}

if (Test-Path -Path "$($Letter):") {
    Write-Host "Error: The drive '$Letter' exists and cannot be used to temporarily mount the drive."
    Exit 1
} 

if (-not (Test-Path -Path $Path)) {
    Write-Host "Error: '$Path' does not exists."
    Exit 1
} 

$virtualDisks = Get-ChildItem -Path $Path -Filter $Filter

$time = Measure-Command {
    [System.Collections.ArrayList]$virtualDiskInfo = @()

    $n = 0
    foreach ($virtualDisk in $virtualDisks) {
        $n++
        $file = $virtualDisk.FullName

        Write-Host "Optimize disk $n of $(($virtualDisks | Measure-Object).Count). File: $file" 
      
        $diskSizeBefore = [math]::Round($(Get-Item -Path $file).length/1MB)

        Write-Host "Create diskpart script to attach vdisk '$file'."
        "select vdisk file=$file" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii
        "attach vdisk" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii -Append

        Start-Sleep -Seconds $PauseTimeAfterEachStep

        Write-Host "Execute diskpart script to attach vdisk '$file'."
        diskpart /s $DiskPartScriptfile

        $partitions = Get-Disk | Where-Object -Property Location -eq $file | Get-Partition        
        foreach ($partition in $partitions){
            $partitionNumber = $partition.PartitionNumber;
            
            Start-Sleep -Seconds $PauseTimeAfterEachStep

            Write-Host "Create diskpart script to assign letter $Letter to partition '$partitionNumber'."
            "select vdisk file=$file" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii
            "select partition $partitionNumber" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii -Append
            "assign letter=$Letter" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii -Append
    
            Start-Sleep -Seconds $PauseTimeAfterEachStep

            Write-Host "Execute diskpart script to assign letter $Letter to partition '$partitionNumber'."
            diskpart /s $DiskPartScriptfile

            Start-Sleep -Seconds $PauseTimeAfterEachStep

            if (Test-Path -Path "$($Letter):") {
                Write-Host "Optimize volume $Letter."
                Optimize-Volume -DriveLetter $Letter -Analyze -Defrag 
            } 

            Write-Host "Create diskpart script to unassign letter $Letter from partition '$partitionNumber'."
            "select vdisk file=$file" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii
            "select partition $partitionNumber" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii -Append
            "remove letter=$Letter" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii -Append
    
            Start-Sleep -Seconds $PauseTimeAfterEachStep

            Write-Host "Execute diskpart script to unassign letter $Letter from partition '$partitionNumber'."
            diskpart /s $DiskPartScriptfile
        }

        Start-Sleep -Seconds $PauseTimeAfterEachStep

        Write-Host "Create diskpart script to detach and shrink vdisk '$file'."
        "select vdisk file=$file" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii
        "detach vdisk" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii -Append
        "compact vdisk" | Out-File -FilePath $DiskPartScriptfile -Encoding ascii -Append

        Start-Sleep -Seconds $PauseTimeAfterEachStep

        Write-Host "Execute diskpart script to detach and shrink vdisk '$file'."
        diskpart /s $DiskPartScriptfile

        $diskSizeAfter = [math]::Round($(Get-Item -Path $file).length/1MB)

        $diskSizeSaving = $diskSizeBefore -  $diskSizeAfter
        Write-Host "Disk size before optimization: $diskSizeBefore MB"
        Write-Host "Disk size after optimization: $diskSizeAfter MB"
        Write-Host "Savings after optimization is: $diskSizeSaving MB"

        $virtualDiskSize = [PSCustomObject]@{
            DiskSizeBefore = $diskSizeBefore
            DiskSizeAfter = $diskSizeAfter
        }

        $virtualDiskInfo.Add($virtualDiskSize) | Out-Null

        Write-Host ""
    }

    $totalDiskSizeBefore = ($virtualDiskInfo.DiskSizeBefore | Measure-Object -Sum).Sum
    $totalDiskSizeAfter = ($virtualDiskInfo.DiskSizeAfter | Measure-Object -Sum).Sum
    $totalDiskSaving = $totalDiskSizeBefore -  $totalDiskSizeAfter
    Write-Host "Total disk size before optimization: $totalDiskSizeBefore MB"
    Write-Host "Total disk size after optimization: $totalDiskSizeAfter MB"
    Write-Host "Total savings on $n disks after optimization is: $totalDiskSaving MB"
}

Write-Host "Optimization runtime was $($time.Minutes) minutes and $($time.Seconds) seconds."

if (Test-Path -Path $DiskPartScriptfile) {
    Remove-Item $DiskPartScriptfile
} 

Exit 0

# SIG # Begin signature block
# MIIojgYJKoZIhvcNAQcCoIIofzCCKHsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA7i8cIeaZyynbt
# pw11WeREMoCmPy5zcxnfieOUbznQ7KCCIaIwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggauMIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMy
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcg
# Q0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXH
# JQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMf
# UBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w
# 1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRk
# tFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYb
# qMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUm
# cJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP6
# 5x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzK
# QtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo
# 80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjB
# Jgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXche
# MBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDig
# NqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd
# 4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiC
# qBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl
# /Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeC
# RK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYT
# gAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/
# a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37
# xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmL
# NriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0
# YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJ
# RyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIG
# wjCCBKqgAwIBAgIQBUSv85SdCDmmv9s/X+VhFjANBgkqhkiG9w0BAQsFADBjMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRp
# Z2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENB
# MB4XDTIzMDcxNDAwMDAwMFoXDTM0MTAxMzIzNTk1OVowSDELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMSAwHgYDVQQDExdEaWdpQ2VydCBUaW1l
# c3RhbXAgMjAyMzCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKNTRYcd
# g45brD5UsyPgz5/X5dLnXaEOCdwvSKOXejsqnGfcYhVYwamTEafNqrJq3RApih5i
# Y2nTWJw1cb86l+uUUI8cIOrHmjsvlmbjaedp/lvD1isgHMGXlLSlUIHyz8sHpjBo
# yoNC2vx/CSSUpIIa2mq62DvKXd4ZGIX7ReoNYWyd/nFexAaaPPDFLnkPG2ZS48jW
# Pl/aQ9OE9dDH9kgtXkV1lnX+3RChG4PBuOZSlbVH13gpOWvgeFmX40QrStWVzu8I
# F+qCZE3/I+PKhu60pCFkcOvV5aDaY7Mu6QXuqvYk9R28mxyyt1/f8O52fTGZZUdV
# nUokL6wrl76f5P17cz4y7lI0+9S769SgLDSb495uZBkHNwGRDxy1Uc2qTGaDiGhi
# u7xBG3gZbeTZD+BYQfvYsSzhUa+0rRUGFOpiCBPTaR58ZE2dD9/O0V6MqqtQFcmz
# yrzXxDtoRKOlO0L9c33u3Qr/eTQQfqZcClhMAD6FaXXHg2TWdc2PEnZWpST618Rr
# IbroHzSYLzrqawGw9/sqhux7UjipmAmhcbJsca8+uG+W1eEQE/5hRwqM/vC2x9XH
# 3mwk8L9CgsqgcT2ckpMEtGlwJw1Pt7U20clfCKRwo+wK8REuZODLIivK8SgTIUlR
# fgZm0zu++uuRONhRB8qUt+JQofM604qDy0B7AgMBAAGjggGLMIIBhzAOBgNVHQ8B
# Af8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAg
# BgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZ
# bU2FL3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFKW27xPn783QZKHVVqllMaPe1eNJ
# MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAG
# CCsGAQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQw
# DQYJKoZIhvcNAQELBQADggIBAIEa1t6gqbWYF7xwjU+KPGic2CX/yyzkzepdIpLs
# jCICqbjPgKjZ5+PF7SaCinEvGN1Ott5s1+FgnCvt7T1IjrhrunxdvcJhN2hJd6Pr
# kKoS1yeF844ektrCQDifXcigLiV4JZ0qBXqEKZi2V3mP2yZWK7Dzp703DNiYdk9W
# uVLCtp04qYHnbUFcjGnRuSvExnvPnPp44pMadqJpddNQ5EQSviANnqlE0PjlSXcI
# WiHFtM+YlRpUurm8wWkZus8W8oM3NG6wQSbd3lqXTzON1I13fXVFoaVYJmoDRd7Z
# ULVQjK9WvUzF4UbFKNOt50MAcN7MmJ4ZiQPq1JE3701S88lgIcRWR+3aEUuMMsOI
# 5ljitts++V+wQtaP4xeR0arAVeOGv6wnLEHQmjNKqDbUuXKWfpd5OEhfysLcPTLf
# ddY2Z1qJ+Panx+VPNTwAvb6cKmx5AdzaROY63jg7B145WPR8czFVoIARyxQMfq68
# /qTreWWqaNYiyjvrmoI1VygWy2nyMpqy0tg6uLFGhmu6F/3Ed2wVbK6rr3M66ElG
# t9V/zLY4wNjsHPW2obhDLN9OTH0eaHDAdwrUAuBcYLso/zjlUlrWrBciI0707NMX
# +1Br/wd3H3GXREHJuEbTbDJ8WC9nR2XlG3O2mflrLAZG70Ee8PBf4NvZrZCARK+A
# EEGKMIIG6DCCBNCgAwIBAgIQd70OBbdZC7YdR2FTHj917TANBgkqhkiG9w0BAQsF
# ADBTMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEpMCcG
# A1UEAxMgR2xvYmFsU2lnbiBDb2RlIFNpZ25pbmcgUm9vdCBSNDUwHhcNMjAwNzI4
# MDAwMDAwWhcNMzAwNzI4MDAwMDAwWjBcMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQ
# R2xvYmFsU2lnbiBudi1zYTEyMDAGA1UEAxMpR2xvYmFsU2lnbiBHQ0MgUjQ1IEVW
# IENvZGVTaWduaW5nIENBIDIwMjAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDLIO+XHrkBMkOgW6mKI/0gXq44EovKLNT/QdgaVdQZU7f9oxfnejlcwPfO
# EaP5pe0B+rW6k++vk9z44rMZTIOwSkRQBHiEEGqk1paQjoH4fKsvtaNXM9JYe5QO
# bQ+lkSYqs4NPcrGKe2SS0PC0VV+WCxHlmrUsshHPJRt9USuYH0mjX/gTnjW4AwLa
# pBMvhUrvxC9wDsHUzDMS7L1AldMRyubNswWcyFPrUtd4TFEBkoLeE/MHjnS6hICf
# 0qQVDuiv6/eJ9t9x8NG+p7JBMyB1zLHV7R0HGcTrJnfyq20Xk0mpt+bDkJzGuOzM
# yXuaXsXFJJNjb34Qi2HPmFWjJKKINvL5n76TLrIGnybADAFWEuGyip8OHtyYiy7P
# 2uKJNKYfJqCornht7KGIFTzC6u632K1hpa9wNqJ5jtwNc8Dx5CyrlOxYBjk2SNY7
# WugiznQOryzxFdrRtJXorNVJbeWv3ZtrYyBdjn47skPYYjqU5c20mLM3GSQScnOr
# BLAJ3IXm1CIE70AqHS5tx2nTbrcBbA3gl6cW5iaLiPcDRIZfYmdMtac3qFXcAzaM
# bs9tNibxDo+wPXHA4TKnguS2MgIyMHy1k8gh/TyI5mlj+O51yYvCq++6Ov3pXr+2
# EfG+8D3KMj5ufd4PfpuVxBKH5xq4Tu4swd+hZegkg8kqwv25UwIDAQABo4IBrTCC
# AakwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwHQYDVR0OBBYEFCWd0PxZCYZjxezzsRM7VxwDkjYRMB8GA1Ud
# IwQYMBaAFB8Av0aACvx4ObeltEPZVlC7zpY7MIGTBggrBgEFBQcBAQSBhjCBgzA5
# BggrBgEFBQcwAYYtaHR0cDovL29jc3AuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25p
# bmdyb290cjQ1MEYGCCsGAQUFBzAChjpodHRwOi8vc2VjdXJlLmdsb2JhbHNpZ24u
# Y29tL2NhY2VydC9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3J0MEEGA1UdHwQ6MDgwNqA0
# oDKGMGh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25pbmdyb290cjQ1
# LmNybDBVBgNVHSAETjBMMEEGCSsGAQQBoDIBAjA0MDIGCCsGAQUFBwIBFiZodHRw
# czovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAHBgVngQwBAzANBgkq
# hkiG9w0BAQsFAAOCAgEAJXWgCck5urehOYkvGJ+r1usdS+iUfA0HaJscne9xthdq
# awJPsz+GRYfMZZtM41gGAiJm1WECxWOP1KLxtl4lC3eW6c1xQDOIKezu86JtvE21
# PgZLyXMzyggULT1M6LC6daZ0LaRYOmwTSfilFQoUloWxamg0JUKvllb0EPokffEr
# csEW4Wvr5qmYxz5a9NAYnf10l4Z3Rio9I30oc4qu7ysbmr9sU6cUnjyHccBejsj7
# 0yqSM+pXTV4HXsrBGKyBLRoh+m7Pl2F733F6Ospj99UwRDcy/rtDhdy6/KbKMxkr
# d23bywXwfl91LqK2vzWqNmPJzmTZvfy8LPNJVgDIEivGJ7s3r1fvxM8eKcT04i3O
# KmHPV+31CkDi9RjWHumQL8rTh1+TikgaER3lN4WfLmZiml6BTpWsVVdD3FOLJX48
# YQ+KC7r1P6bXjvcEVl4hu5/XanGAv5becgPY2CIr8ycWTzjoUUAMrpLvvj1994DG
# TDZXhJWnhBVIMA5SJwiNjqK9IscZyabKDqh6NttqumFfESSVpOKOaO4ZqUmZXtC0
# NL3W+UDHEJcxUjk1KRGHJNPE+6ljy3dI1fpi/CTgBHpO0ORu3s6eOFAm9CFxZdcJ
# JdTJBwB6uMfzd+jF1OJV0NMe9n9S4kmNuRFyDIhEJjNmAUTf5DMOId5iiUgH2vUw
# ggepMIIFkaADAgECAgxGzWBaIDBJGcF2pZgwDQYJKoZIhvcNAQELBQAwXDELMAkG
# A1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExMjAwBgNVBAMTKUds
# b2JhbFNpZ24gR0NDIFI0NSBFViBDb2RlU2lnbmluZyBDQSAyMDIwMB4XDTIzMDYy
# NjEzMDQ1MFoXDTI2MDYyNjEzMDQ1MFowggENMR0wGwYDVQQPDBRQcml2YXRlIE9y
# Z2FuaXphdGlvbjESMBAGA1UEBRMJSFJCIDUwNzgyMRMwEQYLKwYBBAGCNzwCAQMT
# AkRFMRcwFQYLKwYBBAGCNzwCAQITBkhlc3NlbjEiMCAGCysGAQQBgjc8AgEBExFP
# ZmZlbmJhY2ggYW0gTWFpbjELMAkGA1UEBhMCREUxDzANBgNVBAgTBkhlc3NlbjEa
# MBgGA1UEBxMRT2ZmZW5iYWNoIGFtIE1haW4xGjAYBgNVBAkTEUxvZXdlbnN0cmFz
# c2UgNC04MRcwFQYDVQQKEw5EZXZlbG9iaXQgR21iSDEXMBUGA1UEAxMORGV2ZWxv
# Yml0IEdtYkgwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDSeMVExl0l
# dAldYVKpCZXGgIKS3vjULyFUx5r4vHadXGSzkVKGpPhRx3a6Vls26NpWx/xBnY9W
# tcC4eeVFnYEePnSzbjiwOYh/e+D7Gj5JJnA70uvc6jCzRQ6yyZ/TwBm4IZkiAyTy
# U2A6so0btvCmO94RJHTSeJcVSfmPUE5gPW9mDulh9CjPS4xa0UODzRqUksmeu6mu
# tIaZrrB96uAxV3Ls7aPImlzOikEsnpywYZJotz2RWp1BlLiEmSks+h2jd2FgzL7T
# qcokKImCuWa7nTGJtE/RtQ09di3pkISZUiyJSHa2cV4w+Ix8kui7aGOD4SxLvcQA
# mkcJyIyM7F1POJPzqpLbDI8hfDGbiKbWSjsP3KX1yMmPEge+hWUj6IWMTEW/fhA7
# 2DiYvL8K24QjJvEWbBkjPuC161wNTe8T7nRkiEpf76XfpgZQhV08+lSPP/8A8cfJ
# xcY0Ai+zJCmhqRzM6zjJzZDsv6CnpoLrh6QMUnmuleITdJ6igKXLsacqbdPSSEZU
# d1XxGthx1tjtWIIoz09Shbi7FOSUMLfatc6RstiwoY6dcl0IzbCxxA1aKkEOj1ve
# qPebAMPpdHiX5BAEXFrU+KXvBt+1i6xEGfrDwB/94mJlKIs5lp1YrN+fHjadGRhb
# xf8+WsnUsob2K3pKndnukP/IeR+xpWSaeQIDAQABo4IBtjCCAbIwDgYDVR0PAQH/
# BAQDAgeAMIGfBggrBgEFBQcBAQSBkjCBjzBMBggrBgEFBQcwAoZAaHR0cDovL3Nl
# Y3VyZS5nbG9iYWxzaWduLmNvbS9jYWNlcnQvZ3NnY2NyNDVldmNvZGVzaWduY2Ey
# MDIwLmNydDA/BggrBgEFBQcwAYYzaHR0cDovL29jc3AuZ2xvYmFsc2lnbi5jb20v
# Z3NnY2NyNDVldmNvZGVzaWduY2EyMDIwMFUGA1UdIAROMEwwQQYJKwYBBAGgMgEC
# MDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9z
# aXRvcnkvMAcGBWeBDAEDMAkGA1UdEwQCMAAwRwYDVR0fBEAwPjA8oDqgOIY2aHR0
# cDovL2NybC5nbG9iYWxzaWduLmNvbS9nc2djY3I0NWV2Y29kZXNpZ25jYTIwMjAu
# Y3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB8GA1UdIwQYMBaAFCWd0PxZCYZjxezz
# sRM7VxwDkjYRMB0GA1UdDgQWBBQCD4IzXvl28Rah9FD3XtH4SvPIGTANBgkqhkiG
# 9w0BAQsFAAOCAgEAkZHSC1j4fQvZhfZrKSNuVhg7VkvEFyYOCIH6GnnHYaeNRnPh
# FBsKniyMQKqBhxr43iOo0QcbbfXiXl1tjd29451WWwzcq6J4JfanlANjx28cUknB
# 8KpfWvN3r2dwoP2wLo4i8BpWpNi15vGlftMg6W7qSLmEtfCZBA2PVUc/vQpZMYaz
# 1BSY4FUszdA9uncLyirJMrSzQ/qGdjB80wa4E6VePr9PwJ6T6wHPfnu5hrsqgzDG
# cQQGOsXm3VsFeZgEXLwc2G/Y+OHwFjOp1P2Tt1lp3Tzq0WdZINNofq7u6532eyla
# wWeYXvD4wzpWLpTo4Lu0e/aP4p1Nio8fQJaf2xiYd4JoKiAQbEhgnmBScTLDbmaf
# 31OovlXs/u/yYe7XPx7ZUhFwzCQgcL/GeKqPc/9nhj13ZuD4PUyMBLLQAySQA6xC
# ya7yKQDUTK+7IJXh5HCxN1bQffC8b/ZY2w+psoS6aNcjrW/VsEqFEVAcPvfrqx79
# SROwJ5GBA4TH2/2y/XbWbQog95HbS0qIR6Qhym04o+v12JHqLm5NqSBGGlSaopvn
# rXYibIiUwKI5Ltp+Bi0yC0CXtHGlLSVmgC9CQV8H0Axvrb7MvguxqiRGcv5E1Kc4
# rc/BNEDwQVoG3KnR+rydpIJVqb7bRW8ry5v/cymXNXkFnrLuc3ujn21RDC4xggZC
# MIIGPgIBATBsMFwxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52
# LXNhMTIwMAYDVQQDEylHbG9iYWxTaWduIEdDQyBSNDUgRVYgQ29kZVNpZ25pbmcg
# Q0EgMjAyMAIMRs1gWiAwSRnBdqWYMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQB
# gjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIG6WpQrv
# 7rQNlil4vwV74iaebUIZcNKTqB04gaNHjzRPMA0GCSqGSIb3DQEBAQUABIICADDX
# I0PVDRzXkHuBZEpnbn7t+W9DpRzR+gZKHangnQyC4h5rEksS4pPTThZyjz+3H72a
# OUNokhXUdyuPWmVywHrQoGTU34qFKdtRFqiwPzQfaTcOh3DACWxdtxRvqdE3JSPm
# XJqmiq7SuDa5J71VkNgmxnnMJPI+SSFeb4kWL6TR5IxLXIxlEMmUH+WAJWJ6hUPZ
# npBUVauu+V/kqRAcYHEpxvLsnuwB+/8uw8s/Og5Xvi6s6ilUgjaaVgodJ93RHAxR
# AErrFSUxDWbQ0F0WO4h7sweGvUbHgJaDzo5zLMzUIE1dhLeOmKYKR9ka6h2ukziT
# e6RpnttyLgXy54k3FJO/z1v3UfpXhqdHyIhNEFR9NZYESEAkMwRKYBGiXjYIsyoc
# HyBnURYhyTKCsLSFtjHfiOPXe7F+9YC4vaM5Djk62eQ+XVubEFqNw9go5YfAH7OO
# MdMCmC2NS3z75D9P4/RDB36UDMJMQW92CWu8JaUEwqr2Hf6TwT7fixzJvbZLcoE8
# xWoMaoRGT1JHBw3942A0CsVaLckEzS1KVd1HaIi74NE+hhX2f367PHu8jC25n88B
# VaQ6DK1gNA6lYtoIWROZ6Rn83v3j5kBxGZxwzrOJBCZdlA+zwaL2/4h0kYPnCpKo
# QBd78HojJfSWIaBXSf41YLS49RZ5JaX2X3pBjmW3oYIDIDCCAxwGCSqGSIb3DQEJ
# BjGCAw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hB
# MjU2IFRpbWVTdGFtcGluZyBDQQIQBUSv85SdCDmmv9s/X+VhFjANBglghkgBZQME
# AgEFAKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTI0MDgxODE0MjAwNFowLwYJKoZIhvcNAQkEMSIEIGaW39M6WotLn6hrLqMFrSXU
# SOqgcgKJClIIK210FGfjMA0GCSqGSIb3DQEBAQUABIICADcOOQ/0NcymxDmgCxy1
# 9EhHeMJ0u8B2hdEK5z2+i5DIIj3BK4sDNLeYGpwB/WnSYchhDHcwozWmbmdiOTkK
# u7+h501t31rCh5F2ZBROTpuakFvp9rqgndpdNSJz3MHhgdQNjobNA2RFBIfX+zgL
# p7+OQf2ODbAfnyhdFUpIW9PKAAzLvoz0LrSORmIRhGnKrLPG6LeKm2+HVdgQn71p
# f3YhPTTh+/BcdXey9V6we0JBJR4mB9WvJ70iLR3c36K6m2nA9adPRebfZBd4XJFI
# tDWsk2q4U6QEx/wHJ/AxaCwr/ioM+TimD4nkFbxcx+Bz0Hm8hDEXQwjJ7QApDSyu
# 5cYcbwdjynlG/P7V7N+NVA2wj3vJKF4u1W3G+223V6+m5qvVsbN7GYIFZCyXnU9o
# KumLH3TRKDQEekXa35L5NIhSvvlBQ6XKtsBd23NvgDcW1xtexv86T1rLtuMIF0L8
# SJM7eI8iqhuCLKy0FyKT6o6GVmoRWTAa7Y792Ylz550hWVOVb055l3hFWdw/WRH4
# OpbF34HVzbEHCn+wexdAZBMBI5u3sIuYvtgmevaF89FsSypI+jfk/FymAn1L8mG3
# OwS+3bbaOx+bRPLkyq76GCcVSDoPIhioIpUnFzEWYaiim73lJ3NECCJruckPp/p3
# D13m6ZxH2uQqfnwtORp05Qv8
# SIG # End signature block
