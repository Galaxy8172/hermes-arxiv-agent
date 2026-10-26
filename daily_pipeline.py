#!/usr/bin/env python3
"""
daily_pipeline.py —— 每日全链路：抓取 -> LLM 补全 -> 生成数据 -> 提交并推送

供 Windows 计划任务每日调用：
  python daily_pipeline.py

前置：
  - pipeline_config.json（仓库外，含 Kimi key）供 complete_pending.py 使用
  - git 已配置可推送 GitHub（PAT 或 SSH）；未配置时只 commit 不 push，数据仍每日更新在本地

发布门禁：pending_llm_ids.txt 非空时跳过 build_data/push（与上游 publish_viewer.sh 一致），
只发布补全完整的数据。
"""

from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LOG = BASE_DIR.parent / "pipeline_run.log"


def run(title: str, *cmd: str, check: bool = True) -> int:
    print(f"\n{'=' * 56}\n[{datetime.now():%H:%M:%S}] {title}\n{'=' * 56}", flush=True)
    r = subprocess.run(list(cmd), cwd=BASE_DIR)
    if check and r.returncode != 0:
        print(f"[ABORT] {title} 失败（exit={r.returncode}）", flush=True)
    return r.returncode


def main() -> int:
    print(f"\n########## DAILY PIPELINE {datetime.now():%Y-%m-%d %H:%M:%S} ##########", flush=True)

    run("1/4 arXiv 抓取", sys.executable, "monitor.py")

    rc = run("2/4 LLM 补全", sys.executable, "complete_pending.py", check=False)
    if rc != 0:
        print("[WARN] 补全未全部完成，本轮不发布（下轮自动续做）", flush=True)
        return 0

    pending = BASE_DIR / "pending_llm_ids.txt"
    if pending.exists() and pending.read_text(encoding="utf-8").strip():
        print("[WARN] 仍有待补全论文，跳过发布（与上游 publish 门禁一致）", flush=True)
        return 0

    run("3/4 生成 viewer 数据", sys.executable, "viewer" + "\\build_data.py")

    print(f"\n{'=' * 56}\n[4/4] 提交并推送\n{'=' * 56}", flush=True)
    data = BASE_DIR / "viewer" / "papers_data.json"
    if not data.exists():
        print("[WARN] papers_data.json 未生成", flush=True)
        return 0
    run("git add", "git", "add", "viewer/papers_data.json")
    diff = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=BASE_DIR).returncode
    if diff == 0:
        print("[INFO] 数据无变化，无需提交", flush=True)
        return 0
    run("git commit", "git", "commit", "-m", f"data: papers update {datetime.now():%Y-%m-%d}")
    push = run("git push", "git", "push", "origin", "main", check=False)
    if push != 0:
        print("[WARN] push 失败（网络或鉴权）——已本地提交，数据未断；可手动 git push 或稍后重试", flush=True)
    print(f"\n[DONE] {datetime.now():%Y-%m-%d %H:%M:%S}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
