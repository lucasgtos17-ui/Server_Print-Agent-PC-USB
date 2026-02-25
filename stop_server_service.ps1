$ErrorActionPreference = "Stop"

$serviceName = "PrintServerDashboard"
$exePath = Join-Path $PSScriptRoot "dist\PrintServerDashboardService.exe"

Write-Host "Parando servico: $serviceName"

try {
    if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
        if (Test-Path $exePath) {
            & $exePath stop
        } else {
            sc.exe stop $serviceName | Out-Null
        }

        Start-Sleep -Seconds 2
        $svc = Get-Service -Name $serviceName
        Write-Host "Status atual: $($svc.Status)"
    } else {
        Write-Host "Servico '$serviceName' nao encontrado."
    }
} catch {
    Write-Host "Falha ao parar servico: $($_.Exception.Message)"
}

Write-Host "Finalizando processos locais do dashboard (se existirem)..."
try {
    taskkill /IM PrintServerDashboard.exe /F | Out-Null
} catch {
}
try {
    taskkill /IM PrintServerDashboardService.exe /F | Out-Null
} catch {
}

Write-Host "Concluido."
