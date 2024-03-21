<#
    .DESCRIPTION
    Language Setup Part 2
#>

Param (
    [Parameter(Mandatory = $true, HelpMessage = 'Primary Language')]
    [ValidateNotNullorEmpty()] 
    [string] $primaryLanguage,

    [Parameter(Mandatory = $true, HelpMessage = 'Secondary Language')]
    [ValidateNotNullOrEmpty()] 
    [String] $secondaryLanguage,
    
    [Parameter(Mandatory = $true, HelpMessage = 'Primary Input Code')]
    [ValidateNotNullOrEmpty()] 
    [String] $primaryInputCode,
    
    [Parameter(Mandatory = $true, HelpMessage = 'Secondary Input Code')]
    [ValidateNotNullOrEmpty()] 
    [String] $secondaryInputCode,
    
    [Parameter(Mandatory = $true, HelpMessage = 'GeoID')]
    [ValidateNotNullOrEmpty()] 
    [String] $geoID,
    
    [Parameter(Mandatory = $true, HelpMessage = 'TimeZone')]
    [ValidateNotNullOrEmpty()] 
    [String] $timeZone
)

begin {
    function Write-Log {
        [CmdletBinding()]
        <#
            .SYNOPSIS
            Create log function
        #>
        param (
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [System.String] $logPath,
        
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [System.String] $object,
        
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [System.String] $message,
        
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('Information', 'Warning', 'Error', 'Verbose', 'Debug')]
            [System.String] $severity,
        
            [Parameter(Mandatory = $False)]
            [Switch] $toHost
        )
        
        begin {
            $date = (Get-Date).ToLongTimeString()
        }
        process {
            if (($severity -eq "Information") -or ($severity -eq "Warning") -or ($severity -eq "Error") -or ($severity -eq "Verbose" -and $VerbosePreference -ne "SilentlyContinue") -or ($severity -eq "Debug" -and $DebugPreference -ne "SilentlyContinue")) {
                if ($True -eq $toHost) {
                    Write-Host $date -ForegroundColor Cyan -NoNewline
                    Write-Host " - [" -ForegroundColor White -NoNewline
                    Write-Host "$object" -ForegroundColor Yellow -NoNewline
                    Write-Host "] " -ForegroundColor White -NoNewline
                    Write-Host ":: " -ForegroundColor White -NoNewline
        
                    Switch ($severity) {
                        'Information' {
                            Write-Host "$message" -ForegroundColor White
                        }
                        'Warning' {
                            Write-Warning "$message"
                        }
                        'Error' {
                            Write-Host "ERROR: $message" -ForegroundColor Red
                        }
                        'Verbose' {
                            Write-Verbose "$message"
                        }
                        'Debug' {
                            Write-Debug "$message"
                        }
                    }
                }
            }
        
            switch ($severity) {
                "Information" { [int]$type = 1 }
                "Warning" { [int]$type = 2 }
                "Error" { [int]$type = 3 }
                'Verbose' { [int]$type = 2 }
                'Debug' { [int]$type = 2 }
            }
        
            if (!(Test-Path (Split-Path $logPath -Parent))) { New-Item -Path (Split-Path $logPath -Parent) -ItemType Directory -Force | Out-Null }
        
            $content = "<![LOG[$message]LOG]!>" + `
                "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " + `
                "date=`"$(Get-Date -Format "M-d-yyyy")`" " + `
                "component=`"$object`" " + `
                "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + `
                "type=`"$type`" " + `
                "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " + `
                "file=`"`">"
            if (($severity -eq "Information") -or ($severity -eq "Warning") -or ($severity -eq "Error") -or ($severity -eq "Verbose" -and $VerbosePreference -ne "SilentlyContinue") -or ($severity -eq "Debug" -and $DebugPreference -ne "SilentlyContinue")) {
                Add-Content -Path $($logPath + ".log") -Value $content
            }
        }
        end {}
    }

    $logPath = "$env:SYSTEMROOT\TEMP\Deployment_" + (Get-Date -Format 'yyyy-MM-dd')
}

process {
    # Set languages/culture
    try {
        Set-Culture -CultureInfo $primaryLanguage
        Write-Log -Object "LanguageSetup_Part2" -Message "Set Culture to $primaryLanguage" -Severity Information -LogPath $logPath
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($Null -eq $errorMessage) {
            Write-Log -Object "LanguageSetup_Part2" -Message "Failed to set Culture to $primaryLanguage" -Severity Error -LogPath $logPath
        }
        else {
            Write-Log -Object "LanguageSetup_Part2" -Message "$errorMessage" -Severity Error -LogPath $logPath
        }
    }   
                
    # Set UI Language
    try {
        Set-WinUILanguageOverride -Language $primaryLanguage
        Write-Log -Object "LanguageSetup_Part2" -Message "Set UI Language to $primaryLanguage" -Severity Information -LogPath $logPath
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($Null -eq $errorMessage) {
            Write-Log -Object "LanguageSetup_Part2" -Message "Failed to set UI Language to $primaryLanguage" -Severity Error -LogPath $logPath
        }
        else {
            Write-Log -Object "LanguageSetup_Part2" -Message "$errorMessage" -Severity Error -LogPath $logPath
        }
    }
                    
    # Set Location
    try {
        Set-WinHomeLocation -GeoId $geoID
        Write-Log -Object "LanguageSetup_Part2" -Message "Set Windows Home Location to $geoID" -Severity Information -LogPath $logPath
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($Null -eq $errorMessage) {
            Write-Log -Object "LanguageSetup_Part2" -Message "Failed to set Windows Home Location to $geoID" -Severity Error -LogPath $logPath
        }
        else {
            Write-Log -Object "LanguageSetup_Part2" -Message "$errorMessage" -Severity Error -LogPath $logPath
        }
    }
                
    # Set Input Method
    try {
        $newLanguageList = New-WinUserLanguageList -Language "$primaryLanguage"
        $newLanguageList.Add([Microsoft.InternationalSettings.Commands.WinUserLanguage]::new("$secondaryLanguage"))
        $newLanguageList[1].InputMethodTips.Clear()
        $newLanguageList[1].InputMethodTips.Add("$primaryInputCode")
        $newLanguageList[1].InputMethodTips.Add("$secondaryInputCode")
        Set-WinUserLanguageList -LanguageList $newLanguageList -Force
        Write-Log -Object "LanguageSetup_Part2" -Message "Set User Language List" -Severity Information -LogPath $logPath
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($Null -eq $errorMessage) {
            Write-Log -Object "LanguageSetup_Part2" -Message "Failed to set User Language List" -Severity Error -LogPath $logPath
        }
        else {
            Write-Log -Object "LanguageSetup_Part2" -Message "$errorMessage" -Severity Error -LogPath $logPath
        }
    }
                    
    # Set Timezone
    try {
        Set-TimeZone -Name "$timeZone"
        Write-Log -Object "LanguageSetup_Part2" -Message "Set TimeZone to $timeZone" -Severity Information -LogPath $logPath
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($Null -eq $errorMessage) {
            Write-Log -Object "LanguageSetup_Part2" -Message "Failed to set TimeZone to $timeZone" -Severity Error -LogPath $logPath
        }
        else {
            Write-Log -Object "LanguageSetup_Part2" -Message "$errorMessage" -Severity Error -LogPath $logPath
        }
    }

    # Create XML Content
    $XML = @"
<gs:GlobalizationServices xmlns:gs="urn:longhornGlobalizationUnattend">

<!-- user list --> 
<gs:UserList>
<gs:User UserID="Current" CopySettingsToDefaultUserAcct="true" CopySettingsToSystemAcct="true"/> 
</gs:UserList>

<!-- GeoID -->
<gs:LocationPreferences> 
<gs:GeoID Value='$geoID'/>
</gs:LocationPreferences>

<gs:MUILanguagePreferences>
<gs:MUILanguage Value='$primaryLanguage'/>
<gs:MUIFallback Value='$secondaryLanguage'/>
</gs:MUILanguagePreferences>

<!-- system locale -->
<gs:SystemLocale Name='$primaryLanguage'/>

<!-- input preferences -->
<gs:InputPreferences>
<gs:InputLanguageID Action="add" ID='$primaryInputCode' Default="true"/>
<gs:InputLanguageID Action="add" ID='$secondaryInputCode'/>
</gs:InputPreferences>

<!-- user locale -->
<gs:UserLocale>
<gs:Locale Name='$primaryLanguage' SetAsCurrent="true" ResetAllSettings="false"/>
</gs:UserLocale>
</gs:GlobalizationServices>
"@

    # Create XML
    $file = New-Item -Path "$env:SYSTEMROOT\Temp\" -Name "$primaryLanguage.xml" -ItemType File -Value $XML -Force

    # Copy to System and welcome screen
    try {
        Start-Process -FilePath "$env:SYSTEMROOT\System32\Control.exe" -ArgumentList "intl.cpl, , /f:""$($file.Fullname)""" -NoNewWindow -PassThru -Wait | Out-Null
        Write-Log -Object "LanguageSetup_Part2" -Message "Copied settings to System, Welcome Screen and New Users" -Severity Information -LogPath $logPath
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($Null -eq $errorMessage) {
            Write-Log -Object "LanguageSetup_Part2" -Message "Failed to copy settings to System, Welcome Screen and New Users" -Severity Error -LogPath $logPath
        }
        else {
            Write-Log -Object "LanguageSetup_Part2" -Message "$errorMessage" -Severity Error -LogPath $logPath
        }
    }

    # Remove XML
    $file | Remove-Item -Force
}

end {
    # Restart Computer
    Restart-Computer -Force
}