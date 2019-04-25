	& msiexec /i "zabbix_agent-4.2.0-win-amd64-openssl.msi" /qb SERVER=localhost

if (Test-Path -EA Stop C:\Program Files\Zabbix Agent) {
New-Item -EA SilentlyContinue -Path 'C:\Program Files\Zabbix Agent\scripts' -ItemType Directory
    }

if (Test-Path -EA Stop C:\Program Files\Zabbix Agent\scripts){
Copy-Item -Force -Path .\scripts\* -Destination "C:\Program Files\Zabbix Agent\scripts"
    }

$ConfigFile = "C:\Program Files\Zabbix Agent\zabbix_agentd.conf"
$agenthostname=$env:COMPUTERNAME.ToLower()
$proxies="///proxy and server list\\\"


class ZabbixUserParameter
{
    $key
    $value
}

$ReqiredUserParameters = @(
    [ZabbixUserParameter]@{ key = 'script_name_for_zabbix'; value = '_parameters_' }
)

foreach ($r in $ReqiredUserParameters)
{
    $configValue = Get-content $ConfigFile | ? { $_.ToString().Split(',')[0] -eq "UserParameter=$($r.key)" }
    if ( $($configValue | measure).Count -gt 0)
    {
        if ( $configValue.Split(',')[1].Trim() -ne $r.value.Trim() )
        {
            $orig = Get-content $ConfigFile
            $orig | ? { $_ -ne $configValue } | Set-Content $ConfigFile
            "UserParameter=$($r.key),$($r.value)" | Out-File $ConfigFile -Append -encoding ASCII
            $configchange = $true
        }
    } else {
        "UserParameter=$($r.key),$($r.value)" | Out-File $ConfigFile -Append -encoding ASCII
        $configchange = $true
    }
}


if (Get-Content $ConfigFile -Raw '^(Hostname=).*' -notmatch '^(Hostname=$agenthostname).*') {
   (Get-Content $ConfigFile -Raw) -replace '^(Hostname=).*',"`$1$($agenthostname)" | Set-Content $ConfigFile
    $configchange = $true
}

if (Get-Content $ConfigFile -Raw '^(Server=).*' -notmatch '^(Server=$proxies).*') {
    (Get-Content $ConfigFile -Raw) -replace '^(Server=).*',"`$1$($proxies)" | Set-Content $ConfigFile
    $configchange = $true
}

if (Get-Content $ConfigFile -Raw '^(ServerActive=).*' -notmatch '^(ServerActive=$proxies).*') {
    (Get-Content $ConfigFile -Raw) -replace '^(ServerActive=).*',"`$1$($proxies)" | Set-Content $ConfigFile
    $configchange = $true
}

if($configchange -and $((Get-Service -Name 'Zabbix Agent').status -eq 'running') -eq $true) {
    Get-Service -Name 'Zabbix Agent' | Restart-Service
    }
    elseif ($configchange -and $((Get-Service -Name 'Zabbix Agent').status -eq 'stopped') -eq $true) {
    Get-Service -Name 'Zabbix Agent' | Start-Service
    }
