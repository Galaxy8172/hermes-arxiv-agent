#!/usr/bin/env python3
"""
complete_pending.py —— 独立 LLM 补全脚本（替代上游 Hermes agent 的补全步骤）

极简模式（不下载 PDF、不提取单位）：
  - summary_cn：LLM 基于 title/authors/abstract 生成（90-150 字）
  - 结果写回 Excel -> 调 monitor.py --sync-pending-state
  - 小程序端通过 papers_data.json 里的 pdf_url 按需在线打开 PDF

配置从仓库外读取：<repo 的上级目录>/pipeline_config.json
  { "baseURL": "...", "model": "...", "apiKey": "sk-..." }

用法：
  python complete_pending.py            # 补全全部待处理
  python complete_pending.py --limit 5  # 只处理前 5 篇（试跑）
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import time
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR.parent / "data"
CONFIG_PATH = BASE_DIR.parent / "pipeline_config.json"
EXCEL_PATH = DATA_DIR / "papers_record.xlsx"
MONITOR = BASE_DIR / "monitor.py"

HTTP_TIMEOUT = 120
LLM_RETRIES = 2
SAVE_EVERY = 5


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise SystemExit(f"[ERROR] 未找到配置文件：{CONFIG_PATH}（含 baseURL/model/apiKey）")
    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    for k in ("baseURL", "model", "apiKey"):
        if not cfg.get(k):
            raise SystemExit(f"[ERROR] 配置缺少 {k}")
    return cfg


def call_llm(cfg: dict, prompt: str) -> str:
    root = cfg["baseURL"].rstrip().rstrip("/")
    if root.endswith("/v1"):
        root = root[: -len("/v1")]
    url = root + "/v1/messages"
    body = json.dumps({
        "model": cfg["model"],
        "max_tokens": 800,
        "temperature": 0.2,
        "messages": [{"role": "user", "content": prompt}],
    }).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST", headers={
        "Content-Type": "application/json",
        "x-api-key": cfg["apiKey"],
        "anthropic-version": "2023-06-01",
    })
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return "".join(b.get("text", "") for b in data.get("content", []) if isinstance(b, dict))


def extract_json(text: str) -> dict | None:
    m = re.search(r"\{[\s\S]*\}", text.strip())
    if not m:
        return None
    try:
        j = json.loads(m.group(0))
        return j if isinstance(j, dict) else None
    except json.JSONDecodeError:
        return None


def build_prompt(paper: dict) -> str:
    return (
        "你是论文整理助手。根据下面的 arXiv 论文信息，写一段 90-150 字的中文总结，"
        "覆盖：解决什么问题、方法核心思路、主要结果（保留关键数字）。"
        "直接陈述，不要以\"这篇论文\"开头，不要分点，不要出现英文整句。\n"
        "严格只输出 JSON（无代码块标记、无解释）：{\"summary_cn\":\"...\"}\n\n"
        f"标题: {paper.get('title', '')}\n"
        f"作者: {paper.get('authors', '')}\n"
        f"分类: {paper.get('categories', '')}\n"
        f"摘要: {str(paper.get('abstract', ''))[:1200]}"
    )


def load_todo() -> list[dict]:
    wb = load_workbook(EXCEL_PATH, read_only=True)
    ws = wb["Papers"]
    headers = [str(h) if h is not None else "" for h in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {name: i for i, name in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        def g(col: str) -> str:
            v = row[idx[col]] if idx.get(col) is not None and idx[col] < len(row) else None
            return str(v).strip() if v is not None else ""
        if not g("arxiv_id") or g("summary_cn"):
            continue
        rows.append({
            "arxiv_id": g("arxiv_id"), "title": g("title"), "authors": g("authors"),
            "categories": g("categories"), "abstract": g("abstract"),
        })
    wb.close()
    return rows


def update_excel(results: list[dict]) -> int:
    if not results:
        return 0
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Papers"]
    row_by_id = {}
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v:
            row_by_id[str(v).strip()] = r
    updated = 0
    for res in results:
        r = row_by_id.get(res["arxiv_id"])
        if not r:
            continue
        if res.get("summary_cn"):
            ws.cell(row=r, column=8).value = res["summary_cn"]
            updated += 1
    wb.save(EXCEL_PATH)
    return updated


def sync_pending() -> None:
    subprocess.run([sys.executable, str(MONITOR), "--sync-pending-state"], cwd=BASE_DIR, check=False)


def main() -> None:
    limit = 0
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])

    cfg = load_config()
    todo = load_todo()
    if limit:
        todo = todo[:limit]
    print(f"[INFO] 待补全 {len(todo)} 篇 | 模型 {cfg['model']}", flush=True)
    if not todo:
        sync_pending()
        print("[OK] 无待补全论文")
        return

    done, failed = [], []
    for i, paper in enumerate(todo, 1):
        print(f"[{i}/{len(todo)}] {paper['arxiv_id']} | {paper['title'][:60]}", flush=True)
        res = None
        for attempt in range(1, LLM_RETRIES + 1):
            try:
                raw = call_llm(cfg, build_prompt(paper))
                j = extract_json(raw)
                if j and j.get("summary_cn"):
                    res = {"arxiv_id": paper["arxiv_id"], "summary_cn": str(j["summary_cn"]).strip()}
                    break
                print(f"    [WARN] 第 {attempt} 次返回无法解析：{raw[:80]}", flush=True)
            except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
                print(f"    [WARN] 第 {attempt} 次请求失败：{e}", flush=True)
                time.sleep(3 * attempt)
        if res:
            done.append(res)
            print(f"    [OK] 摘要 {len(res['summary_cn'])} 字", flush=True)
        else:
            failed.append(paper["arxiv_id"])
            print("    [FAIL] 放弃本篇，留待下轮", flush=True)

        if len(done) and len(done) % SAVE_EVERY == 0:
            print(f"    [SAVE] 写回 {update_excel(done[-SAVE_EVERY:])} 条", flush=True)

    if len(done) % SAVE_EVERY:
        print(f"[SAVE] 写回 {update_excel(done[len(done) - (len(done) % SAVE_EVERY):])} 条", flush=True)

    sync_pending()
    tail = ('（失败: ' + ', '.join(failed[:5]) + ('…' if len(failed) > 5 else '') + '）') if failed else ''
    print(f"[DONE] {datetime.now().isoformat()} 成功 {len(done)}/{len(todo)}" + tail, flush=True)


if __name__ == "__main__":
    main()
