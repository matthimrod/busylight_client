$Config = Get-Content -Raw ~\.dotfiles\busylight.json | ConvertFrom-Json
$Host.UI.RawUI.WindowTitle = $Config.WindowTitle

function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

$LastActivity = ""

if ((Get-NetConnectionProfile).Name -match $Config.MyNetworkName) {
    try {
        while($true) {
            $LogContent = Get-Content $Config.TeamsLogFile -tail 2000 | Select-String -Pattern 'StatusIndicatorStateService\: Added (\w+) [^|]*'
            if ($LogContent -and $LogContent.Matches.Length -ge 1) {
                $activity = $LogContent.Matches[$LogContent.Matches.Length - 1].Groups[1].Value
    
                if ($null -ne $activity -and $activity -ne $LastActivity) {
                    $retries = $Config.MaxRetry
                    do {
                        Write-Output "$(Get-TimeStamp) Setting status to $activity"
                        try {
                            $Global:ProgressPreference = 'SilentlyContinue'
                            $null = Invoke-RestMethod -ProgressAction SilentlyContinue -Uri $Config.URL -Method 'Post' -Body @{ state = $activity }
                            $LastActivity = $activity
                            $retries = 0
                        } catch [System.Object] {
                            $retries--
                            Start-Sleep -Seconds $Config.RetryWait
                        }
                    } until ($retries -eq 0) 
                }
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
