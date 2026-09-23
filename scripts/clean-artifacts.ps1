param(
    # 保留最新的 N 个 artifact，其余删除（默认 1）
    [int]$Keep = 1
)

# 清理旧 artifacts：避免 Actions 存储被历史制品占满（私库有配额，占满会导致上传失败）。
# 需要 permissions: actions: write。任何失败都只告警、不阻断构建。
$repo = $env:GITHUB_REPOSITORY
$headers = @{
    Authorization          = "Bearer $env:GH_TOKEN"
    Accept                 = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
}

# 分页取全部 artifacts
$arts = @()
$page = 1
while ($true) {
    try {
        $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/actions/artifacts?per_page=100&page=$page" -Headers $headers
    } catch {
        Write-Warning "查询 artifacts 失败（跳过清理）：$_"
        exit 0
    }
    $batch = @($resp.artifacts)
    $arts += $batch
    if ($batch.Count -lt 100 -or $page -ge 20) { break }
    $page++
}

Write-Host "现有 artifacts：$($arts.Count) 个（保留最新 $Keep 个）"
if ($arts.Count -le $Keep) {
    Write-Host "无需清理"
    exit 0
}

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
