param(
    # 保留最新的 N 个 artifact，其余删除（默认 1）
    [int]$Keep = 1,
    # >0 时同时清理旧的 workflow run（保留最新 N 条运行记录及其日志；0 = 不清理）
    [int]$KeepRuns = 0
)

# 清理旧 artifacts / 旧运行记录：避免 Actions 存储被历史制品与日志占满
# （私库有配额，占满会导致上传失败）。需要 permissions: actions: write。
# 任何失败都只告警、不阻断构建。
$repo = $env:GITHUB_REPOSITORY
$headers = @{
    Authorization          = "Bearer $env:GH_TOKEN"
    Accept                 = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
}

# ---- 1) 清理旧 workflow run（保留最新 $KeepRuns 条；进行中的一律不删） ----
if ($KeepRuns -gt 0) {
    $runs = @()
    $page = 1
    $ok = $true
    while ($true) {
        try {
            $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/actions/runs?per_page=100&page=$page" -Headers $headers
        } catch {
            Write-Warning "查询运行记录失败（跳过运行记录清理）：$_"
            $ok = $false
            break
        }
        $batch = @($resp.workflow_runs)
        $runs += $batch
        if ($batch.Count -lt 100 -or $page -ge 20) { break }
        $page++
    }
    if ($ok) {
        Write-Host "现有运行记录：$($runs.Count) 条（保留最新 $KeepRuns 条）"
        # 安全阀：最新一条运行必须已经产出制品（上传成功），否则不删任何运行记录，
        # 避免上传失败时把上一次构建留下的唯一可用制品一起删掉。
        $newestRun = $runs | Sort-Object -Property created_at -Descending | Select-Object -First 1
        try {
            $latest = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/actions/artifacts?per_page=1" -Headers $headers
            $latestArtifacts = @($latest.artifacts)
        } catch {
            $latestArtifacts = @()
        }
        $latestRunId = if ($latestArtifacts.Count -gt 0) { $latestArtifacts[0].workflow_run.id } else { 0 }
        if ($null -eq $newestRun -or $latestRunId -ne $newestRun.id) {
            Write-Warning "最新运行尚未产出制品（可能上传失败）：本次跳过运行记录清理，保留历史制品"
        }
        elseif ($runs.Count -gt $KeepRuns) {
            $keepIds = @($runs | Sort-Object -Property created_at -Descending |
                         Select-Object -First $KeepRuns | ForEach-Object { $_.id })
            $dropRuns = @($runs | Where-Object { $keepIds -notcontains $_.id -and $_.status -eq 'completed' })
            foreach ($r in $dropRuns) {
                try {
                    Invoke-RestMethod -Method Delete -Uri "https://api.github.com/repos/$repo/actions/runs/$($r.id)" -Headers $headers | Out-Null
                    Write-Host "已删除运行记录：#$($r.run_number) id=$($r.id)（$($r.head_branch) / $($r.created_at)）"
                } catch {
                    Write-Warning "删除运行记录 id=$($r.id) 失败：$_"
                }
            }
        }
    }
}

# ---- 2) 清理旧 artifacts（保留最新 $Keep 个） ----
$arts = @()
$page = 1
$ok = $true
while ($true) {
    try {
        $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/actions/artifacts?per_page=100&page=$page" -Headers $headers
    } catch {
        Write-Warning "查询 artifacts 失败（跳过清理）：$_"
        $ok = $false
        break
    }
    $batch = @($resp.artifacts)
    $arts += $batch
    if ($batch.Count -lt 100 -or $page -ge 20) { break }
    $page++
}
if (-not $ok) { exit 0 }

Write-Host "现有 artifacts：$($arts.Count) 个（保留最新 $Keep 个）"
if ($arts.Count -gt $Keep) {
    $sorted = $arts | Sort-Object -Property created_at -Descending
    $drop = $sorted | Select-Object -Skip $Keep
    foreach ($a in $drop) {
        try {
            Invoke-RestMethod -Method Delete -Uri "https://api.github.com/repos/$repo/actions/artifacts/$($a.id)" -Headers $headers | Out-Null
            Write-Host "已删除 artifact：$($a.name) id=$($a.id) 创建于 $($a.created_at)"
        } catch {
            Write-Warning "删除 artifact id=$($a.id) 失败：$_"
        }
    }
}
