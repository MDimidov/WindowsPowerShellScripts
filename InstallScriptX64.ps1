Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# # Запазва текущите настройки за захранване
# $currentPowerScheme = powercfg /getactivescheme

# Задава временен план за захранване, който предотвратява заспиването и изключването на дисплея
powercfg -change -monitor-timeout-ac 0
powercfg -change -monitor-timeout-dc 0
powercfg -change -standby-timeout-ac 0
powercfg -change -standby-timeout-dc 0

Write-Host "The script begin" -ForegroundColor Green


# # Remove default Windows Photos Viewer
# Get-AppxPackage *photo* | Remove-AppxPackage

# ----------Working on windows 10 and later-----------
# Disable real-time protection
Set-MpPreference -DisableRealtimeMonitoring $true
Write-Host "Real-time protection is disabled" -ForegroundColor Yellow

# Add current folder for exclusion
$folderPath = "$pwd"
Add-MpPreference -ExclusionPath $folderPath
# ----------------------------------------------------

# Disable SmartScreen
Write-Host "Disabling SmartScreen..." -ForegroundColor Yellow
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Value "Off"

# Define the paths or URLs to the installers and the display names used for checking installation
$programs = @(
    @{ Name = "7-Zip"; Path = "$pwd\x64\7z2408-x64.exe"; Args = "/S"; Check = "7-Zip" },
    @{ Name = "Win Rar"; Path = "$pwd\x64\WinRAR v6.10 Final (x64) EN + BG\winrar-x64-610.exe"; Args = "/S"; Check = "winrar" },
    @{ Name = "Adobe Photoshop"; Path = "$pwd\x64\Photoshop\Adobe Photoshop 2020 21.2.4.323 (x64) RePack by SanLex.exe"; Args = "/silent /norestart"; Check = "Photoshop" },
    @{ Name = "Adobe Acrobat"; Path = "$pwd\x32\Adobe.Acrobat.Pro.v2022.003.20310x86\Adobe.Acrobat.Pro.v2022.003.20310x86.exe"; Args = "-y -nr -gm2"; Check = "Acrobat" },
    @{ Name = "Google Chrome"; Path = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"; Args = "/silent /install"; Check = "Google Chrome" },
    @{ Name = "Brave"; Path = "https://laptop-updates.brave.com/download/BRV013?bitness=64"; Args = "/silent /install"; Check = "Brave" },
    @{ Name = "4K Video Downloader"; Path = "https://dl.4kdownload.com/app/4kvideodownloaderplus_1.9.0_x64_online.exe?source=website&_gl=1*98nit9*_ga*NTIwMzY0ODU2LjE3MjY3NDUwNDU.*_ga_KRN0PZWTN9*MTcyNjc0NTA0NS4xLjEuMTcyNjc0NTY2Ni4wLjAuMA.."; Args = "/S"; Check = "4K Video" },
    @{ Name = "Viber"; Path = "$pwd/x64/ViberSetup.exe"; Args = "/S"; Check = "Viber" },
    @{ Name = "Daum PotPlayer"; Path = "$pwd\x64\Daum PotPlayer v220106 [1.7.21589] Stable (x32-x64) + Skins\PotPlayerSetup64.exe"; Args = "/S /install /norestart"; Check = "PotPlayer" },
    @{ Name = "K-Lite Codec"; Path = "$pwd\K-Lite Codec Pack 16.1.2 Mega_Full_Standard_Basic\K-Lite_Codec_Pack_1612_Standard.exe"; Args = "/silent /install /norestart"; Check = "K-Lite" },
    @{ Name = "Microsoft Office 2019 LTS"; Path = "$pwd\x64\Microsoft Office LTSC Pofessional Plus 2021 v16.0.14332.20204  (x86 & x64)\Tools for Office2021\setup.exe"; Args = "/configure `"$pwd\x64\MS Office\Configuration.xml`""; Check = "Microsoft Office" },
    @{ Name = "Favorite games"; Path = "$pwd\x32\favorite-games_aff123.exe"; Args = "/silent /install"; Check = "favorite-games" },
    @{ Name = "Winamp"; Path = "$pwd\x32\WINAMp\Winamp Pro 5.666 Build 3516 Final\winamp5666_full_all.exe"; Args = "/S"; Check = "winamp" },
    @{ Name = "JPEGView"; Path = "msiexec.exe"; Args = "/i `"$pwd\x64\JPEGView64_en-us_1.3.46.msi`" /quiet"; Check = "JPEGView" },
    @{ Name = "uTorrent"; Path = "$pwd\uTorrent_3.5.5.45966.exe"; Args = "/silent"; Check = "torrent" }
)

# Function to check if a program is installed
function IsProgramInstalled($programName) {
    $uninstallKeyPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Office\ClickToRun\Configuration\*"
    )

    foreach ($path in $uninstallKeyPaths) {

        if (Test-Path $path) {

            $installedPrograms = Get-ItemProperty -Path $path | Select-Object DisplayName

            foreach ($program in $installedPrograms.DisplayName) {

                if ($program -like "*$programName*") {

                    return $true

                }
            }
        } 

        # else {
        #     Write-Warning "Path $path does not exist."
        # }

    }

    return $false
}

foreach ($program in $programs) {
    if (IsProgramInstalled $program.Check) {
        Write-Host "$($program.Name) is already installed. Skipping..." -ForegroundColor Cyan
    }
    else {
        Write-Output "Installing $($program.Name)..."

        if ($program.Path -like "http*") {
            # Download installer if it's a URL
            $installerPath = "$env:TEMP\$($program.Name.Replace(' ', '')).exe"
            Invoke-WebRequest -Uri $program.Path -OutFile $installerPath
        }
        else {
            $installerPath = $program.Path
        }

        # Run the installer
        try{
            Start-Process -FilePath $installerPath -ArgumentList $program.Args -NoNewWindow -Wait

            # If the program is Adobe Acrobat, run additional commands
            if ($program.Name -eq "Adobe Acrobat") {
                $acrobatHelperPath = "C:\Users\$Env:UserName\AppData\Local\Temp\AdobeAcrobat\AcrobatHelper.exe" 
                if (Test-Path $acrobatHelperPath) {
                    Start-Process -FilePath $acrobatHelperPath -ArgumentList "/S /AUTO" -Wait
                }
                else {
                    Write-Host "AcrobatHelper.exe not found in $acrobatHelperPath" -ForegroundColor Yellow
                }
                
                # Remove the temporary installation directory
                $acrobatTempDir = "C:\Users\$Env:UserName\AppData\Local\Temp\AdobeAcrobat"
                if (Test-Path $acrobatTempDir) {
                    Remove-Item -Path $acrobatTempDir -Recurse -Force
                }
                else {
                    Write-Host "The folder $acrobatTempDir does not exist." -ForegroundColor Yellow
                }
            }

            if(IsProgramInstalled $program.Check){
                Write-Host "$($program.Name) installation completed." -ForegroundColor Green
            } else {
                Write-Host "$($program.Name) instalation failed" -ForegroundColor Red
            }
        }
        catch{
           Write-Host "$($program.Name) instalation failed" -ForegroundColor Red
        }
    }
}






# Определете пътя до работния плот
$desktopPath = [System.Environment]::GetFolderPath('Desktop')

# Define the paths or URLs to the installers and the display names used for checking installation
$shortcuts = @(
    # @{ Name = "7-Zip"; Path = "$pwd\x64\7z2408-x64.exe"; Args = "/S"; Check = "7-Zip" },
    @{ Name = "Adobe Photoshop"; TargetPath = "C:\Program Files\Adobe Photoshop 2020\Photoshop.exe"; Directory = "C:\Program Files\Adobe Photoshop 2020"; Check = "Photoshop"; Action = "Create" },
    @{ Name = "Favorite games"; TargetPath = "C:\Program Files (x86)\Favorite-Games\Favorite-Games.exe"; Directory = "C:\Program Files (x86)\Favorite-Games"; Check = "favorite-games"; Action = "Create" },
    @{ Name = "This PC"; TargetPath = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"; Directory = "C:\Windows\System32\imageres.dll, -109"; Check = "favorite-games"; Action = "Create" },
    @{ Name = "JPEGView"; TargetPath = "JPEGView.lnk"; Directory = "C:\Program Files\JPEGView"; Check = "JPEGView"; Action = "Delete" },
    @{ Name = "PotPlayer"; TargetPath = "PotPlayer 64 bit.lnk"; Directory = "C:\Program Files\DAUM\PotPlayer"; Check = "PotPlayer"; Action = "Delete" }

)

foreach($shcut in $shortcuts){
    if (IsProgramInstalled $shcut.Check) {
        if ($shcut.Action -eq "Create"){
            # Създаване на пряк път на работния плот
            $WScriptShell = New-Object -ComObject WScript.Shell
            
            # Път към програмата, която искате да свържете
            $targetPath = $shcut.TargetPath  # Поправете пътя според това, къде е инсталирана играта
            
            # Определете пътя до прекия път
            $shortcutPath = "$desktopPath\$($shcut.Name).lnk"
            
            # Създайте пряк път
            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $targetPath  # Път към .exe файла на програмата
            $shortcut.WorkingDirectory = $shcut.Directory  # Работната директория (трябва да е същата като инсталацията)
            $shortcut.WindowStyle = 1  # Нормален прозорец
            # $shortcut.IconLocation = "$targetPath, 0"  # Икона (по подразбиране от самия .exe файл)
            $shortcut.Save()
            
            Write-Host "$($shcut.Name) created shortcut" -ForegroundColor Green
        } elseif ($shcut.Action -eq "Delete"){
            # Delete shortcut if exist
            if(Test-Path -Path "$($desktopPath)\$($shcut.TargetPath)"){
                Remove-Item -Path "$($desktopPath)\$($shcut.TargetPath)" -Force
                Write-Host "$($shcut.Name) deleted shortcut" -ForegroundColor Yellow
            } elseif (Test-Path -Path "C:\Users\Public\Desktop\$($shcut.TargetPath)") {
                Remove-Item -Path "C:\Users\Public\Desktop\$($shcut.TargetPath)" -Force
                Write-Host "$($shcut.Name) deleted shortcut" -ForegroundColor Yellow
            }
        }
    }
}

# Set skin to winamp S7-Reflex
Start-Process "C:\Program Files (x86)\Winamp\winamp.exe" -Wait
Start-Process "$pwd\x32\WINAMp\S7Reflex.wal" -Wait

# -------------------------------------------------------------------
# Define the paths for copying text to the hosts file
$sourceFile = "$pwd\ads.txt"
$hostsFile = "C:\Windows\System32\drivers\etc\hosts"

# Read the content of both files
$sourceContent = Get-Content -Path $sourceFile
$hostsContent = Get-Content -Path $hostsFile
$isHostContainsSource = $true

# Loop through each line in the source file and check if it's in the hosts file
foreach ($line in $sourceContent) {
    if (-not ($hostsContent -contains $line)) {
        $isHostContainsSource = $false
        break
    }
}

if (-not $isHostContainsSource) {
    # Check if running as administrator
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You do not have Administrator rights to make changes to the hosts file."
        return
    }

    # Copy the content from the source file and append it to the hosts file
    Get-Content $sourceFile | Add-Content $hostsFile
    Write-Host "Text has been successfully added to the hosts file." -ForegroundColor Green
}
else {
    Write-Host "Text already exists in the hosts file. Skipping..." -ForegroundColor Yellow
}
# -------------------------------------------------------------------


# Copy files from one directory to another
try {
    Copy-Item -Path "$pwd\x64\WinRAR v6.10 Final (x64) EN + BG\rarreg.key" -Destination "C:\Program Files\WinRAR\" -Force
    Write-Host "WinRar has been successfully activated."  -ForegroundColor Green
}
catch {
    Write-Error "WinRar throw an error during activation."
}

# Copy files from one directory to another / translate WinRar to BG
try {
    Copy-Item -Path "$pwd\x64\WinRAR v6.10 Final (x64) EN + BG\BG\*" -Destination "C:\Program Files\WinRAR\" -Force
    Write-Host "WinRar has been successfully translated to BG." -ForegroundColor Green
}
catch {
    Write-Error "WinRar failed to translate."
}

# Copy files from one directory to another
try {
    Copy-Item -Path "$pwd\JPEGView.ini" -Destination "C:\Program Files\JPEGView\" -Force
    Write-Host "JPEGView setting updated."  -ForegroundColor Green
}
catch {
    Write-Error "JPEGView setting update failed."
}

## ----------------------------------------------------------------
$extensions = @(".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".psd", ".webp", ".NEF", ".CR2")
$defaultProgram = "C:\Program Files\JPEGView\JPEGView.exe"

foreach ($ext in $extensions) {
    # Променяме регистъра за файловите разширения
    $classKey = "HKCU:\Software\Classes\$ext"
    if (-not (Test-Path $classKey)) {
        New-Item -Path $classKey
    }
    
    # Задаваме програмата по подразбиране
    Set-ItemProperty -Path $classKey -Name "(Default)" -Value "jpegviewfile"
    
    # Настройваме командата за отваряне на файла
    $commandKey = "HKCU:\Software\Classes\jpegviewfile\shell\open\command"
    if (-not (Test-Path $commandKey)) {
        New-Item -Path $commandKey -Force
    }
    Set-ItemProperty -Path $commandKey -Name "(Default)" -Value "`"$defaultProgram`" `"%1`""
}

# Обновяване на настройките за файловите асоциации
# Stop-Process -Name explorer -Force
# Start-Process explorer
Write-Host "JPEGView set as default image app." -ForegroundColor Green
## ----------------------------------------------------------------


# Activate windows and office

try {
    Start-Process -FilePath "$pwd\MAS_AIO.cmd" -Wait
} catch {
    Invoke-RestMethod "https://get.activated.win" | Invoke-Expression
} 
    
# Проверка на активацията на Windows
$windowsActivationStatus = (Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object { $_.Name -like "Windows*" -and $_.PartialProductKey }).LicenseStatus

if ($windowsActivationStatus -eq 1) {
    Write-Host "Windows is activated." -ForegroundColor Green
} else {
    Write-Host "Windows is not activated." -ForegroundColor Red
}


# Path to ospp.vbs script
$osppPath = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"  # For 64-bit Office
# $osppPath = "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs"  # For 32-bit Office

# Check Office activation status
$officeActivationStatus = cscript.exe $osppPath /dstatus | Out-String

# Check if Office is activated
if ($officeActivationStatus -match "LICENSE STATUS:  ---LICENSED---") {
    Write-Host "Microsoft Office is activated." -ForegroundColor Green
} else {
    Write-Host "Microsoft Office is not activated." -ForegroundColor Red
}


#-----------------------------------------------------
# Re-enable SmartScreen
Write-Host "Re-enabling SmartScreen..." -ForegroundColor Cyan
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Value "RequireAdmin"


# Връща настройките обратно след изпълнението на скрипта
powercfg -change -monitor-timeout-ac 10
powercfg -change -monitor-timeout-dc 10
powercfg -change -standby-timeout-ac 30
powercfg -change -standby-timeout-dc 30

# Enable real-time protection
Set-MpPreference -DisableRealtimeMonitoring $true
Write-Host "Real-time protection is enabled" -ForegroundColor Green

# #-----------------------------------------------------------
# Set system locale to Bulgarian
Set-WinSystemLocale bg-BG

# Set the UI language to Bulgarian
Set-WinUILanguageOverride -Language bg-BG

# # Set the user language list to Bulgarian
# # Set-WinUserLanguageList -LanguageList bg-BG -Force

# Set the culture (region) to Bulgaria
Set-Culture bg-BG

# Set the home location (geo region) to Bulgaria
Set-WinHomeLocation -GeoId 35

Write-Host "Region and language settings updated to Bulgaria. Please restart your computer for all changes to take effect." -ForegroundColor Green

$answer = Read-Host "Do you want to restart your PC? Press [Y] for Yes"

if ($answer -eq 'Y') {
    Write-Host "Restarting the PC..." -ForegroundColor Green
    Restart-Computer
} else {
    Write-Host "You chose not to restart the PC." -ForegroundColor Yellow
}
# #-----------------------------------------------------------

Write-Host "The script is finished" -ForegroundColor Green