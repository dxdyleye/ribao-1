#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""清理 GitHub Actions 存储：artifacts + 旧运行记录（保留最新），可选重跑失败的构建。

删除 artifact / workflow run 都会释放 Actions 存储配额（配额统计 = artifacts + 运行日志）。

用法：
  GH_TOKEN=xxx python3 清理_actions记录.py                      # 清理 artifacts(留1) + 旧runs(留1)
  GH_TOKEN=xxx python3 清理_actions记录.py --dry-run            # 只看不删
  GH_TOKEN=xxx python3 清理_actions记录.py --keep-runs 3        # 保留最近 3 条运行
  GH_TOKEN=xxx python3 清理_actions记录.py --rerun-failed       # 额外重跑最新的失败运行
  GH_TOKEN=xxx python3 清理_actions记录.py --repo OWNER/REPO

Token 权限（细粒度）：Repository access 必须包含目标仓库（私有仓库需选 All repositories
或 Only select repositories 并勾选该仓库）；Permissions → Actions: Read and write。
（经典 Token：勾选 repo + workflow）
"""
import argparse
import json
import os
import urllib.error
import urllib.request

API = 'https://api.github.com'


def api(url, token, method='GET', body=None):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, method=method, data=data)
    req.add_header('Authorization', 'Bearer %s' % token)
    req.add_header('Accept', 'application/vnd.github+json')
    req.add_header('X-GitHub-Api-Version', '2022-11-28')
    req.add_header('User-Agent', 'actions-cleaner')
    if data:
        req.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            raw = r.read().decode()
            return r.status, (json.loads(raw) if raw.strip() else None)
    except urllib.error.HTTPError as e:
        return e.code, {'error': e.read().decode('utf-8', 'replace')[:300]}


def paged(url_fmt, token, key):
    out, page = [], 1
    while True:
        st, d = api(url_fmt % page, token)
        if st != 200:
            raise SystemExit('查询失败（HTTP %s）：%s' % (st, d))
        batch = d.get(key, [])
        out.extend(batch)
        if len(batch) < 100 or page >= 20:
            return out
        page += 1


def mb(n):
    return '%.1f MB' % (n / 1024.0 / 1024.0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo', default='dxdyleye/ribao-1')
    ap.add_argument('--keep-artifacts', type=int, default=1, help='保留最新 N 个 artifact（默认 1）')
    ap.add_argument('--keep-runs', type=int, default=1, help='保留最新 N 条运行记录（默认 1）')
    ap.add_argument('--no-runs', action='store_true', help='不删除运行记录（只清 artifacts）')
    ap.add_argument('--rerun-failed', action='store_true', help='清理后重跑最新一次失败的运行')
    ap.add_argument('--dry-run', action='store_true')
    args = ap.parse_args()

    token = os.environ.get('GH_TOKEN') or os.environ.get('GITHUB_TOKEN')
    if not token:
        raise SystemExit('请设置 GH_TOKEN 环境变量')

    st, me = api('%s/user' % API, token)
    if st != 200:
        raise SystemExit('Token 无效（HTTP %s）：%s' % (st, me))
    print('已认证：%s' % me.get('login'))

    R = args.repo
    st, repo = api('%s/repos/%s' % (API, R), token)
    if st != 200:
        raise SystemExit('无法访问仓库 %s（HTTP %s）——请检查 Token 的 Repository access 是否包含该仓库'
                         '（私有仓库需选 All repositories 或 Only select repositories 并勾选它）：%s'
                         % (R, st, repo))
    print('仓库：%s（private=%s）' % (repo.get('full_name'), repo.get('private')))

    # ---------- artifacts ----------
    arts = paged('%s/repos/%s/actions/artifacts?per_page=100&page=%%d' % (API, R), token, 'artifacts')
    total = sum(a['size_in_bytes'] for a in arts)
    print('\n=== artifacts：%d 个，合计 %s ===' % (len(arts), mb(total)))
    for a in sorted(arts, key=lambda x: x['created_at'], reverse=True):
        print('   id=%-12s %-34s %10s created=%s expired=%s' % (
            a['id'], a['name'], mb(a['size_in_bytes']), a['created_at'], a['expired']))

    newest = sorted(arts, key=lambda x: x['created_at'], reverse=True)[:args.keep_artifacts]
    keep_ids = {a['id'] for a in newest}
    drop = [a for a in arts if a['id'] not in keep_ids]
    freed = 0
    if drop:
        print('\n删除 %d 个旧 artifact（保留最新 %d 个）：' % (len(drop), args.keep_artifacts))
        for a in drop:
            if args.dry_run:
                print('   [dry-run] 将删除 id=%s %s %s' % (a['id'], a['name'], mb(a['size_in_bytes'])))
                freed += a['size_in_bytes']
                continue
            st, d = api('%s/repos/%s/actions/artifacts/%s' % (API, R, a['id']), token, 'DELETE')
            if st in (204, 200):
                freed += a['size_in_bytes']
                print('   已删除 id=%s %s（%s）' % (a['id'], a['name'], mb(a['size_in_bytes'])))
            else:
                print('   删除失败 id=%s（HTTP %s）：%s' % (a['id'], st, d))
    else:
        print('\n无需删除 artifact。')

    # ---------- workflow runs ----------
    dropped_runs = 0
    runs = []
    if not args.no_runs:
        runs = paged('%s/repos/%s/actions/runs?per_page=100&page=%%d' % (API, R), token, 'workflow_runs')
        print('\n=== 运行记录：%d 条（保留最新 %d 条，其余删除，连带清掉其日志）==='
              % (len(runs), args.keep_runs))
        for r in runs[:args.keep_runs]:
            print('   保留 run=%-12s #%-4s %-9s %s' % (r['id'], r['run_number'], r['conclusion'], r['created_at']))
        for r in runs[args.keep_runs:]:
            if args.dry_run:
                print('   [dry-run] 将删除 run=%s #%s %s' % (r['id'], r['run_number'], r['created_at']))
                dropped_runs += 1
                continue
            st, d = api('%s/repos/%s/actions/runs/%s' % (API, R, r['id']), token, 'DELETE')
            if st in (204, 200):
                dropped_runs += 1
                print('   已删除 run=%s #%s（%s）' % (r['id'], r['run_number'], r['created_at']))
            else:
                print('   删除失败 run=%s（HTTP %s）：%s' % (r['id'], st, d))

    print('\n=== 汇总 ===')
    print('删除 artifact %d 个（释放约 %s）；删除运行记录 %d 条。%s'
          % (len(drop), mb(freed), dropped_runs, '（dry-run，未实际删除）' if args.dry_run else ''))

    # ---------- 重跑失败的构建 ----------
    if args.rerun_failed and not args.dry_run:
        if not runs:
            runs = paged('%s/repos/%s/actions/runs?per_page=20&page=%%d' % (API, R), token, 'workflow_runs')
        target = next((r for r in runs if r['conclusion'] in ('failure', 'cancelled', 'timed_out')), None)
        if target is None:
            print('\n没有需要重跑的失败运行。')
        else:
            st, d = api('%s/repos/%s/actions/runs/%s/rerun-failed-jobs' % (API, R, target['id']), token, 'POST')
            print('\n重跑 run=%s #%s -> HTTP %s %s' % (
                target['id'], target['run_number'], st, '' if st in (201, 202) else d))


if __name__ == '__main__':
    main()
