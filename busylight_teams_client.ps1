$Config = Get-Content -Raw ~\.dotfiles\busylight.json | ConvertFrom-Json
$Host.UI.RawUI.WindowTitle = $Config.WindowTitle

function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

$LastActivity = ""

if ((Get-NetConnectionProfile).Name -match $Config.MyNetworkName) {
    while($true) {
        $LogContent = Get-Content $Config.TeamsLogFile -tail 1000 | Select-String -Pattern 'StatusIndicatorStateService\: Added (\w+) [^|]*'
        if ($null -ne $activity) {
            $activity = $LogContent.Matches[$LogContent.Matches.Length - 1].Groups[1].Value

            if ($activity -ne $LastActivity) {
                $retries = $Config.MaxRetry
                do {
                    Write-Output "$(Get-TimeStamp) Setting status to $activity"
                    try {
                        $result = Invoke-RestMethod -Uri $Config.URL -Method 'Post' -Body @{ state = $activity }
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
} else {
    Write-Output "Not connected to home network!"
}
