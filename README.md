# Optimize-Vhdx.ps1

Powershell script that attempts to shrink vhdx files.

Execute the script with administrator rights.

This script tries to shrink vhdx files. It defragments the file and then tries to shrink the file by the unused disk space via diskpart.

The script has three phases.

1. It creates and runs a diskpart script to attach the vdisk.
2. It attach each partition and defragments the volume using the Powershell Optimize-Volume commandlet.
3. It creates a diskpart script to detach and shrink the disk.

The script works without the Powershell commandlet Optimize-VHD and don't need the Hyper-V Management Tools.

The script is inspired by Johan Arwidmark https://www.deploymentresearch.com/optimizing-vhdx-files-in-a-hyper-v-lab/

# MIT License

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

# Mandatory parameter

## Path

The path where the vhdx files are located.

# Optional parameter

## DiskPartScriptfile

The file name for the temporary diskpart script. Default file name is 'diskpart.txt'.

## Filter

The file filter to be used. \*.vhdx for all vhdx files in the directory. sample.vhdx for a specific vhdx file. Default filter is '\*.vhdx'.

## Letter

Drive letter that is used temporarily to mount the vdisk. Default letter is 'T'.

## PauseTimeAfterEachStep

The time (in seconds) to be paused after each processing step. Default time are 3 seconds

# Examples

.\Optimize-Vhdx.ps1 -Path c:\UserProfiles

Change drive letter.
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -Letter J

Change file name of diskpart script.
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -DiskPartScriptfile script.txt

Change pause time after each step.
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -PauseTimeAfterEachStep 15

Optimize specific vhdx files by custom file filter.
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -Filter UVHD-S-1-5-21-\*.vhdx

Optimize specific vhdx file.
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -Filter UVHD-S-1-5-21-2967749299-2196513631-28783610-1237.vhdx

All parameters combined.
.\Optimize-Vhdx.ps1 -Path c:\UserProfiles -DiskPartScriptfile script.txt -Filter UVHD-S-1-5-21-2967749299-2196513631-28783610-1237.vhdx -Letter J -PauseTimeAfterEachStep 15
