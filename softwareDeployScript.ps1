
#Programs: 7-Zip,Adobe Reader, Adobe Air, Adobe FlashPlayer IE, Adobe FlashPlayer Chrome, Adobe Shockwave , 
#          CutePDF Writer, Chrome, JDK6, JDK7, JDK8, JDK9, JRE6, JRE7, JRE8, JRE9, .NET, Silverlight, PSv4, Symantec Cloud
#TODO: Add CMD Status Checking
#      Write to Log File on share if Error
#      
$path = $env:temp
$computer = $env:COMPUTERNAME
$ErrorActionPreference = "Stop"
$start_time = Get-Date
$empty_line = ""
# Determine the architecture of a machine                                                     # Credit: Tobias Weltner: "PowerTips Monthly vol 8 January 2014"
If ([IntPtr]::Size -eq 8) {
    $empty_line | Out-String
    "Running in a 64-bit subsystem" | Out-String
    $64 = $true
    $bit_number = "64"
    $registry_paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $empty_line | Out-String
} Else {
    $empty_line | Out-String
    "Running in a 32-bit subsystem" | Out-String
    $64 = $false
    $bit_number = "32"
    $registry_paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )    
    $empty_line | Out-String
} # else
$net_paths = @(
    "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP",
    "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
)
$base_path = "\\SERVER\SHARE\pdq_pack\repository"

$software = @{
    _7zip=@{
            name='7-Zip';
            version=@('16.04');
            Script64="7-Zip x64.bat";
            Script32="7-Zip x86.bat";
            Dir64="\7-zip\x64\";
            Dir32="\7-zip\x86\";
        };
    AdobeReader=@{
            name='Adobe Acrobat Reader';
            version=@('15.023.20053', '17.012.20098');
            Script64="Adobe Acrobat Reader DC.bat";
            Script32="Adobe Acrobat Reader DC.bat";
            Dir64="\adobe\acrobat_reader_dc\x86\";
            Dir32="\adobe\acrobat_reader_dc\x86\";
        }; 
    AdobeAir=@{
            name='Adobe AIR';
            version=@('27.0.0.124');
            Script64="Adobe AIR.bat";
            Script32="Adobe AIR.bat";
            Dir64="\adobe\air\x86\";
            Dir32="\adobe\air\x86\";
        };
    AdobeFlashChrome=@{
            name='PPAPI';
            version=@('27.0.0.130');
            Script64="Adobe Flash Player (Chrome).bat";
            Script32="Adobe Flash Player (Chrome).bat";
            Dir64="\adobe\flash_player\chrome\";
            Dir32="\adobe\flash_player\chrome\";
        };
    AdobeFlashIE=@{
            name='Adobe Flash Player 27 ActiveX';
            version=@('27.0.0.130', '26.0.0.151');
            Script64="Adobe Flash Player (IE).bat";
            Script32="Adobe Flash Player (IE).bat";
            Dir64="\adobe\flash_player\internet_explorer\";
            Dir32="\adobe\flash_player\internet_explorer\";
        };
    AdobeShockwave=@{
            name='Adobe Shockwave';
            version=@('12.2.9.199');
            Script64="Adobe Shockwave.bat";
            Script32="Adobe Shockwave.bat";
            Dir64="\adobe\shockwave\x86\";
            Dir32="\adobe\shockwave\x86\";
        };
    CutePDF=@{
            name='CutePDF';
            version=@('3.0');
            Script64="CutePDF Writer v3.0 x86.bat";
            Script32="CutePDF Writer v3.0 x86.bat";
            Dir64="\cutepdf\cutepdf\x86\";
            Dir32="\cutepdf\cutepdf\x86\";
        };
    GoogleChrome=@{
            name='Google Chrome';
            version=@('61.0.3163.100');
            Script64="googlechromestandaloneenterprise.bat";
            Script32="googlechromestandaloneenterprise x86.bat";
            Dir64="\google\chrome_enterprise\x64\";
            Dir32="\google\chrome_enterprise\x86\";
        };
    JRE7=@{
            name='Java 7';
            version=@('7.0.800');
            Script64="jre-7-x64.bat";
            Script32="jre-7-i586.bat";
            Dir64="\java\jre\7\x64\";
            Dir32="\java\jre\7\x86\";
        };
    JRE8=@{
            name='Java 8';
            version=@('8.0.1440.1');
            Script64="jre-8-x64.bat";
            Script32="jre-8-i586.bat";
            Dir64="\java\jre\8\x64\";
            Dir32="\java\jre\8\x86\";
        };
    <#JRE9=@{
            name='Java 9';
            version=@('9.0.0.0');
            Script64="\java\jre\9\x64\jre-9-x64.bat";
            Script32="";
        };#>
    NET35=@{
            name='.NET Framework 3.5';
            version=@('3.5');
            Script64=".NET_Framework_v3.5.bat";
            Script32=".NET_Framework_v3.5.bat";
            Dir64="\microsoft\dot-net-framework\";
            Dir32="\microsoft\dot-net-framework\";
        };
    NET4=@{
            name='.NET Framework 4.0';
            version=@('4.0');
            Script64=".NET_Framework_v4.0.bat";
            Script32=".NET_Framework_v4.0.bat";
            Dir64="\microsoft\dot-net-framework\";
            Dir32="\microsoft\dot-net-framework\";
        };
    NET45=@{
            name='.NET Framework 4.5';
            version=@('4.5','4.7');
            Script64=".NET_Framework_v4.5.bat";
            Script32=".NET_Framework_v4.5.bat";
            Dir32="\microsoft\dot-net-framework\";
            Dir64="\microsoft\dot-net-framework\";
        };
    Silverlight=@{
            name='Silverlight';
            version=@('5.1.50901');
            Script64="Silverlight x64.bat";
            Script32="Silverlight x86.bat";
            Dir64="\microsoft\silverlight\x64\";
            Dir32="\microsoft\silverlight\x86\";
        };
    PowerShell4=@{
            name='Powershell';
            version=@('4.0');
            Script64="install_ps4.bat";
            Script32="install_ps4.bat";
            Dir64="\powershell\";
            Dir32="\powershell\";
        };
    SymantecCloud=@{
            name='Symantec.cloud';
            version=@('3.00.10.2737');
            Script64="symantec.bat";
            Script32="symantec.bat";
            Dir64="\symantec\";
            Dir32="\symantec\";
        };
    RenameCutePDF=@{
            name='Rename Cute PDF';
            version=@('0.0');
            Script64="rename_cutepdf.bat";
            Script32="rename_cutepdf.bat";
            Dir64="\utilities\rename_cutepdf\";
            Dir32="\utilities\rename_cutepdf\";
        };
    EmptyRecycleBin=@{
            name='Empty Recycle Bin';
            version=@('0.0');
            Script64="empty_all_recycle_bins.bat";
            Script32="empty_all_recycle_bins.bat";
            Dir64="\utilities\empty_all_recycle_bins\";
            Dir32="\utilities\empty_all_recycle_bins\";
        };

}


Function Check-InstalledSoftware ($display_name, $display_version) {
    Write-Host "Checking if ${display_name} is installed..."

    #get current package with name
    $current = Get-ItemProperty $registry_paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -Like "*${display_name}*"}

    #if package found then check versions
    If($current){
        Write-Host "$display_name is Installed, Checking Versions..."
        #get version on machine
        $currentVersion = $current.DisplayVersion
        Write-Host "${display_name} is installed with version(s): ${currentVersion}"

        #for each new version check if current version matches exactly
        foreach ($new_version in $display_version) {
            Write-Host "Checking version: ${new_version}"
            #if new version found on machine...
            If(Get-ItemProperty $registry_paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -Like "*${display_name}*" -and $_.DisplayVersion -Like "*${new_version}*" }){
                Write-Host "${display_name} is already installed with version(s): $new_version"
                return $true
            }
        }

        #if it doesnt match then check if highest version on machine is higher than new version
        $max_new_version = ($display_version | Sort-Object | Select -last 1)
        $max_installed_version = ($currentVersion | Sort-Object -Descending | Select -last 1)
        Write-Host "Max New Version $max_new_version, Max Installed Version $max_installed_version"
        If($max_installed_version -ge $max_new_version){
            Write-Host "Installed Version (${max_installed_version}) is greater than New Version(${max_new_version}) so no need to install"
            return $true
        }

    #if not installed        
    }Else{
        Write-Host "${display_name} is Not in Installed List"


        #if NET then check NET versions in registry
        If($display_name.StartsWith(".NET")){
            Write-Host "Checking NET Registry Path..."
            foreach ($new_version in $display_version) {
                Write-Host "Checking version: ${new_version}"
                #if new version found on machine...
                If(Get-ChildItem $net_paths | ForEach-Object {Get-ItemProperty $_.pspath} | Where-Object {$_.Version -Like "*${new_version}*" -or $_.PSChildName -Like "*${new_version}*"}){
                    Write-Host "${display_name} is already installed with version(s): $new_version"
                    return $true
                }
            }
        }

        #if Powershell then check Powershell versions in registry
        If($display_name.StartsWith("Powershell")){
            Write-Host "Checking Powershell Registry Path..."
            foreach ($new_version in $display_version) {
                Write-Host "Checking version: ${new_version}"
                #if new version found on machine...
                If($PSVersionTable.PSVersion.Major -ge [decimal]$new_version){
                    Write-Host "${display_name} is already installed with version(s): $new_version"
                    return $true
                }
            }
        }

    }

    Return Get-ItemProperty $registry_paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -Like "*${display_name}*" -and $_.DisplayVersion -Like "*${display_version}*" }
}

#go through dictionary above one app at a time
foreach ($h in $software.Keys) {
    $name = $software.Item($h).Item("name")
    $version=$software.Item($h).Item("version")
    Write-Host "Current Software Installing...: ${h} with Version: ${version}"
    
    #NEW Path will be in C:\Windows\GPO\...
    $new_path = $ENV:systemroot + "\GPO\$h"
    New-Item -ItemType Directory -Force -Path $new_path

    #get source path and script name
    If($64){
        $source_path = $base_path + $software.Item($h).Item("Dir64")
        $script = Join-Path $new_path $software.Item($h).Item("Script64")
    }Else{
        $source_path = $base_path + $software.Item($h).Item("Dir32")
        $script = Join-Path $new_path $software.Item($h).Item("Script32")
    }

    #sync to corresponding folder in C:\Windows\GPO 
    Write-Host "Copying from $source_path to $new_path"
    Robocopy $source_path $new_path /MIR /E /R:3 /W:3 /NJH /NJS
    Write-Host "Done Copying"

    Write-Host "Starting Check-InstalledSoftware with parameters $name and $version"
    $output = Check-InstalledSoftware $name $version
    Write-Host "Return from Check-InstalledSoftware with output '$output'"
    If($output){
        Write-Host "- ${h} already installed"
    }Else{
        #if not installed then execute script
        Write-Host "+ Will install ${h}"
        If($script){
            Write-Host "Executing ${script}"
            cmd.exe /c $script
        }Else{
            Write-Host "Script not available for current platform"
        }
    }
    Write-Host "Done Installing ${h}"
    #Start-Sleep -s 3
    Write-Host ""
}
