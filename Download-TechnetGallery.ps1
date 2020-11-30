<#

.SYNOPSIS

This script downloads the files of defined contributions from the Microsoft Technet Gallery.
Those who are downloading either sealed or unsealed Management Packs for SCOM are able to import them automatically.

(c) 29.08.2017, Patrick Seidl, s2 - seidl solutions

.DESCRIPTION

Some notes to the script:
1. Sometimes the script does not download any file. Close your open browsers; Try the script again.
2. The script does not work in my ISE session here; just plain PowerShell.
3. Edit the URL list to whatever you'd like to grab.
4. Set autoDownload and autoImport to false if you'd like to be prompted.
5. The script downloads into C:\Temp by default. $env:userprofile\Downloads does not work if you have to use another account to import into SCOM.

#>

# enter the URL of each gallery post you'd like to download (each in a new line)
$gallerylinks = @("https://gallery.technet.microsoft.com/PoSh-Technet-Files-229c3559"
"https://gallery.technet.microsoft.com/MP-SCOM-Trace-Helper-Tasks-70380ca4"
"https://gallery.technet.microsoft.com/MP-SCOM-HealthService-63c99fcf"
"https://gallery.technet.microsoft.com/PoSh-Reset-Monitor-On-c288374a"
"https://gallery.technet.microsoft.com/PSGW-Connected-Users-20817b14"
"https://gallery.technet.microsoft.com/PSGW-Authorized-Users-e566c5aa"
"https://gallery.technet.microsoft.com/PSGW-Computers-without-f4199ba7"
"https://gallery.technet.microsoft.com/PoSh-Set-Resource-Pool-aea4e7be"
"https://gallery.technet.microsoft.com/PoSh-Show-Resource-Pool-40d9b18f"
"https://gallery.technet.microsoft.com/PoSh-Set-Gateway-Primary-cbb9c77b"
"https://gallery.technet.microsoft.com/PSGW-Gateway-Server-MS-7b2df7bf"
"https://gallery.technet.microsoft.com/SCOM-Agent-Management-b96680d5"
"https://gallery.technet.microsoft.com/SQL-Server-RunAs-Addendum-0c183c32")

# set to false if you don't want to automatically download and/or import into SCOM
$autoDownload = $true
$autoImport = $true

#$folderPath = Join-Path $env:userprofile "Downloads"
$folderpath = "C:\Temp"
New-Item -path $folderpath -Name "Technet_Gallery" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
$folderPath = Join-Path $folderPath "Technet_Gallery"

""
Write-Host "Links to process:"
$gallerylinks 

foreach ($gallerylink in $gallerylinks) {
    "-"*70
    # get the site
    try {
        ""
        "Download Web Site"
        $site = Invoke-WebRequest -Uri $gallerylink
    } catch {
        Write-Host "Could not find the site requested" -ForegroundColor Red
        Continue
    }

    # get the download link
    $download = $site.links | ? {$_.outerHTML -match $gallerylink.Split("/")[-1] -and $_.class -eq "button"}
    $link = "https://gallery.technet.microsoft.com" + ($download).'data-url'

    ""
    Write-Host "Title:      " $site.ParsedHtml.title
    Write-Host "File Name:  " $download.outerText
    Write-Host "Last Update:" ($site.AllElements | ? {$_.Id -eq "LastUpdated"}).outertext
    
    ""
    if ($autoDownload -eq $false) {
        Write-Host "Do you want to download the file? [y/n]" -ForegroundColor Yellow
        $getfile = Read-Host
    } else {
        $getfile = "y"
    }
    if ($getfile -eq "y") {
        New-Item -path $folderpath -Name $download.outerText.Remove($download.outerText.Length-($download.outerText.split(".")[-1].length+1),($download.outerText.split(".")[-1].length+1)) -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        $path = Join-Path $folderPath $download.outerText.Remove($download.outerText.Length-($download.outerText.split(".")[-1].length+1),($download.outerText.split(".")[-1].length+1))

        # download the file
        ""
        "Download File"
        Invoke-WebRequest $link -OutFile (Join-Path $path $download.outerText)

        # unzip
        if ($download.outerText -like "*.zip") {
            ""
            "Unzip File"
            $shell = new-object -com shell.application
            $zip = $shell.NameSpace((Join-Path $path $download.outerText))
            foreach($item in $zip.items()) {
                Remove-Item (Join-Path $path $item.Name) -Force -ErrorAction SilentlyContinue
                $shell.Namespace($path).copyhere($item)
            }
            Remove-Item (Join-Path $path $download.outerText) -Force -ErrorAction SilentlyContinue
        }

        if ($download.outerText -like "*.mp" -or $download.outerText -like "*.xml" -or $item.name -like "*.mp" -or $item.name -like "*.xml") {
            # import to your SCOM
            ""
            if ($autoImport -eq $false) {
                Write-Host "Do you want to import the MP to your SCOM? [y/n]" -ForegroundColor Yellow
                $importMp = Read-Host
            } else {
                $importMp = "y"
            }
            if ($importMP -eq "y") {
                try {
                    "Load SCOM Module"
                    Import-Module OperationsManager
                    "Import MP"
                    $allMP = Get-ChildItem $path
                    "Importing: "
                    $allMP.Name
                    Import-SCOMManagementPack -Fullname $allMP.FullName
                } catch {
                    if (!$cred) {
                        ""
                        Write-Host "Please use SCOM admin account" -ForegroundColor Red
                        $cred = Get-Credential
                    }
                    if (!$sdkMs) {
                        ""
                        Write-Host "Enter the name of your Management Server [default: localhost]" -ForegroundColor Yellow
                        $sdkMs = Read-Host
                        if (!$sdkMs) {$sdkMs = "localhost"}
                    }
                    try { 
                        Start-Job -Name MpImport -Credential $cred -ArgumentList @($path, $sdkMs) -ScriptBlock {
                            $path = $args[0]
                            $sdkMs = $args[1]
                            "Load SCOM Module"
                            Import-Module OperationsManager
                            New-SCManagementGroupConnection -ComputerName $sdkMs
                            "Import MP"
                            $allMP = Get-ChildItem $path
                            "Importing: "
                            $allMP.Name
                            Import-SCOMManagementPack -Fullname $allMP.FullName
                        } | Out-Null
                        Wait-Job -Name MpImport | Out-Null
                        # Receive-Job -Name MpImport | fl *
                    } catch {
                        Write-Host "Please check if the working dir for PowerShell is set to C:\" -ForegroundColor Red
                        Break
                    }
                }
            }
        $item = $null
        }
    }
}