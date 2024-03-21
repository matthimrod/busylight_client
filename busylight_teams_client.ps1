$Config = Get-Content -Raw ~\.dotfiles\busylight.json | ConvertFrom-Json
$Host.UI.RawUI.WindowTitle = $Config.WindowTitle

function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

function Get-WebcamInUse {
    $webcam_last_used = Get-ChildItem HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam `
       | ForEach-Object { if( $_.Property.Contains('LastUsedTimeStop') ) { $_.GetValue('LastUsedTimeStop') } }
    return $webcam_last_used -Contains 0
}

function Get-MicrophoneInUse {
    $microphone_last_used = Get-ChildItem HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone `
       | ForEach-Object { if( $_.Property.Contains('LastUsedTimeStop') ) { $_.GetValue('LastUsedTimeStop') } }
    return $microphone_last_used -Contains 0
}

$LastActivity = ""

if ((Get-NetConnectionProfile).Name -match $Config.MyNetworkName) {
    try {
        while($true) {
            $webcam = Get-WebcamInUse
            $microphone = Get-MicrophoneInUse

            if ($webcam -or $microphone) {
                $activity = 'on-the-phone'
            } else {
                $activity = 'away'
            }
    
            if ($null -ne $activity -and $activity -ne $LastActivity) {
                $retries = $Config.MaxRetry
                do {
                    Write-Output "$(Get-TimeStamp) Setting status to $activity"
                    try {
                        $null = Invoke-RestMethod -ProgressAction SilentlyContinue -Uri $Config.URL -Method 'Post' -Body @{ state = $activity }
                        $LastActivity = $activity
                        $retries = 0
                    } catch [System.Object] {
                        $retries--
                        Start-Sleep -Seconds $Config.RetryWait
                    }
                } until ($retries -eq 0) 
            }
            Start-Sleep –Seconds $Config.PollingInterval 
        }
    }
    finally {
        # If something happens that terminates the script, try to turn off the light before we goooooo.....
        $null = Invoke-RestMethod -Uri $Config.URL -Method 'Post' -Body @{ state = 'off' }
    }
} else {
    Write-Output "Not connected to home network!"
}
