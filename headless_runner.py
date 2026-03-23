#!/usr/bin/env python3
"""
headless_runner.py
==================
NotebookLM × YouTube 每日自動分析 – 無 GUI 版本
用於 GitHub Actions / 任何支援 Python 3.11+ 的無頭環境

環境變數（設定在 GitHub Secrets）：
  YOUTUBE_API_KEY      YouTube Data API Key
  NOTEBOOK_ID          NotebookLM Notebook ID（若策略為 single_notebook）
  ANALYSIS_COMMAND     分析指令，例如：台股 美股 ETF｜請整理...
  TOP_N                搜尋影片數（預設 10）
  REGION_CODE          地區代碼（預設 TW）
  LANGUAGE             語言代碼（預設 zh-Hant）
  MIN_MINUTES          最短片長分鐘（預設 0）
  STORAGE_STATE_B64    notebooklm storage_state.json 的 base64 編碼內容
  EMAIL_TO             收件人 Email（逗號分隔多人）
  EMAIL_FROM           寄件人 Gmail 帳號
  EMAIL_APP_PASSWORD   Gmail App Password（非一般密碼）
  LINE_NOTIFY_TOKEN    LINE Notify Token
  GDRIVE_SA_JSON_B64   Google Drive Service Account JSON 的 base64 內容
  GDRIVE_FOLDER_ID     Google Drive 目標資料夾 ID
  EXTRA_TICKERS        額外要查詢報價的代號（逗號分隔，例如：2330,AAPL,GC=F）
"""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import os
import re
import smtplib
import sys
import traceback
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import Callable, Dict, List, Optional
from urllib.parse import urlparse

import requests

# ── Optional deps (safe import) ─────────────────────────────────────────────
try:
    from notebooklm import NotebookLMClient
    try:
        from notebooklm import ReportFormat, InfographicOrientation, InfographicDetail
    except Exception:
        ReportFormat = InfographicOrientation = InfographicDetail = None
    _HAS_NOTEBOOKLM = True
except Exception:
    _HAS_NOTEBOOKLM = False

try:
    import yfinance as yf
    _HAS_YFINANCE = True
except Exception:
    _HAS_YFINANCE = False

try:
    from googleapiclient.discovery import build as _gdrive_build
    from googleapiclient.http import MediaFileUpload
    from google.oauth2 import service_account
    _HAS_GDRIVE = True
except Exception:
    _HAS_GDRIVE = False

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("headless")


# =============================================================================
# Config
# =============================================================================
class Config:
    def __init__(self):
        self.youtube_api_key    = os.environ.get("YOUTUBE_API_KEY", "").strip()
        self.notebook_id        = os.environ.get("NOTEBOOK_ID", "").strip()
        self.analysis_command = os.environ.get(
            "ANALYSIS_COMMAND",
            "台股 美股 ETF 投資 財經｜"
            "請整理今天最熱門的投資影片，重點萃取："
            "1. 提到的股票代號、ETF代號、目標價、支撐壓力位 "
            "2. 博主觀點與市場共識分歧 "
            "3. 每支影片的核心論點（1-2句）"
            "4. 播放量特別高的原因（標題策略、發布時機、獨家資訊）"
            "輸出格式：先給5-8點摘要重點，再給來源影片表格，最後給股票/ETF彙整表"
        ).strip()
        self.top_n              = int(os.environ.get("TOP_N", "10") or 10)
        self.region_code        = os.environ.get("REGION_CODE", "TW").strip()
        self.language           = os.environ.get("LANGUAGE", "zh-Hant").strip()
        self.min_minutes        = int(os.environ.get("MIN_MINUTES", "0") or 0)
        self.storage_state_b64  = os.environ.get("STORAGE_STATE_B64", "").strip()
        self.email_to           = [e.strip() for e in
                                   os.environ.get("EMAIL_TO", "").split(",") if e.strip()]
        self.email_from         = os.environ.get("EMAIL_FROM", "").strip()
        self.email_app_password = os.environ.get("EMAIL_APP_PASSWORD", "").strip()
        self.line_token         = os.environ.get("LINE_NOTIFY_TOKEN", "").strip()
        self.gdrive_sa_b64      = os.environ.get("GDRIVE_SA_JSON_B64", "").strip()
        self.gdrive_folder_id   = os.environ.get("GDRIVE_FOLDER_ID", "").strip()
        self.extra_tickers      = [t.strip() for t in
                                   os.environ.get("EXTRA_TICKERS", "").split(",") if t.strip()]

    def validate(self) -> List[str]:
        errors = []
        if not self.youtube_api_key:
            errors.append("YOUTUBE_API_KEY 未設定")
        if not self.notebook_id and _HAS_NOTEBOOKLM:
            errors.append("NOTEBOOK_ID 未設定")
        if not self.storage_state_b64 and _HAS_NOTEBOOKLM:
            errors.append("STORAGE_STATE_B64 未設定（需要 NotebookLM 登入資料）")
        return errors


# =============================================================================
# Utility functions (copied from main GUI, no dependencies)
# =============================================================================
@dataclass
class VideoItem:
    rank: int
    title: str
    channel: str
    url: str
    published_at: str
    published_local: str
    views_text: str
    views_num: int
    duration_text: str
    duration_seconds: int
    description: str = ""


def parse_views_to_int(v) -> int:
    try:
        return int(v)
    except Exception:
        return 0


def format_views(n: int) -> str:
    return f"{n:,}"


def iso8601_duration_to_seconds(text: str) -> int:
    if not text:
        return 0
    pattern = re.compile(
        r"P(?:(?P<days>\d+)D)?(?:T(?:(?P<hours>\d+)H)?(?:(?P<minutes>\d+)M)?(?:(?P<seconds>\d+)S)?)?"
    )
    m = pattern.fullmatch(text)
    if not m:
        return 0
    return (int(m.group("days") or 0) * 86400 +
            int(m.group("hours") or 0) * 3600 +
            int(m.group("minutes") or 0) * 60 +
            int(m.group("seconds") or 0))


def seconds_to_hms(seconds: int) -> str:
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def to_taipei_display(utc_iso: str) -> str:
    if not utc_iso:
        return ""
    try:
        dt = datetime.fromisoformat(utc_iso.replace("Z", "+00:00"))
        tw = dt.astimezone(timezone(timedelta(hours=8)))
        return tw.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return utc_iso


def safe_str(v) -> str:
    return str(v) if v is not None else ""


# =============================================================================
# Smart keyword extraction (from v37)
# =============================================================================
_KEYWORD_MAP: Dict[str, List[str]] = {
    "台股": ["台股 分析", "台股 盤勢"],
    "美股": ["美股 分析", "美股 投資"],
    "etf": ["ETF 分析", "ETF 投資 台灣"],
    "股票": ["股票 分析", "選股策略"],
    "投資": ["投資 理財", "投資策略"],
    "財經": ["財經 新聞", "財經 分析"],
    "ai": ["AI 科技股", "人工智慧 投資"],
    "人工智慧": ["AI 科技股", "人工智慧 投資"],
    "半導體": ["半導體 股票", "半導體 ETF"],
    "科技股": ["科技股 分析", "科技 ETF"],
    "黃金": ["黃金 投資", "黃金 走勢"],
    "原油": ["原油 分析", "能源 投資"],
    "比特幣": ["比特幣 分析", "加密貨幣 投資"],
    "fed": ["Fed 利率 分析", "聯準會 政策"],
    "利率": ["利率 投資影響", "Fed 升降息"],
    "通膨": ["通膨 投資策略", "CPI 分析"],
    "台積電": ["台積電 分析", "TSMC 股價"],
    "輝達": ["輝達 NVIDIA 分析", "NVDA 股價"],
    "存股": ["存股 策略", "高股息 存股"],
    "美元": ["美元指數 分析", "匯率 投資"],
    "債券": ["債券 投資", "美國公債"],
}


def infer_queries(command: str) -> List[str]:
    cmd = command.strip()
    queries: List[str] = []
    seen: set = set()

    def _add(q: str):
        q = q.strip()
        if q and q not in seen:
            seen.add(q)
            queries.append(q)

    # Format A: "關鍵字1 關鍵字2｜分析要求"
    for sep in ["｜", "|"]:
        if sep in cmd:
            kw_part = cmd.split(sep, 1)[0].strip()
            raw_kws = re.split(r"[\s,，、]+", kw_part)
            raw_kws = [k.strip() for k in raw_kws if k.strip() and len(k.strip()) >= 2]
            if raw_kws:
                for kw in raw_kws:
                    mapped = False
                    for key, terms in _KEYWORD_MAP.items():
                        if kw.lower() == key.lower() or key.lower() in kw.lower():
                            _add(terms[0])
                            mapped = True
                            break
                    if not mapped and len(kw) >= 2:
                        _add(f"{kw} 投資" if "投資" not in kw and "分析" not in kw else kw)
                if queries:
                    return queries[:6]
            break

    # Format B: free-form
    cmd_lower = cmd.lower()
    for key, terms in _KEYWORD_MAP.items():
        if key.lower() in cmd_lower:
            _add(terms[0])

    return queries[:6] if queries else ["台股 分析", "美股 投資", "ETF 分析", "財經 今日"]


# =============================================================================
# YouTube Searcher (headless version)
# =============================================================================
class YouTubeSearcher:
    SEARCH_URL = "https://www.googleapis.com/youtube/v3/search"
    VIDEOS_URL = "https://www.googleapis.com/youtube/v3/videos"

    def __init__(self, api_key: str, logger: Callable[[str], None] = log.info):
        self.api_key = api_key.strip()
        self.logger  = logger
        self.session = requests.Session()

    def _request(self, url: str, params: dict) -> dict:
        p = dict(params)
        p["key"] = self.api_key
        resp = self.session.get(url, params=p, timeout=40)
        if resp.status_code != 200:
            try:
                detail = resp.json()
            except Exception:
                detail = resp.text
            # Friendly quota error
            reason = ""
            if isinstance(detail, dict):
                errors = detail.get("error", {}).get("errors", [{}])
                reason = (errors[0] if errors else {}).get("reason", "")
            if reason in {"quotaExceeded", "dailyLimitExceeded"} or resp.status_code == 403:
                raise RuntimeError(
                    f"YouTube API 配額已用完（{resp.status_code}）。"
                    f"台北時間明天 08:00 後重置。原因：{reason}"
                )
            raise RuntimeError(f"YouTube API 失敗 ({resp.status_code}): {detail}")
        return resp.json()

    def search_videos(
        self,
        command: str,
        top_n: int = 10,
        region_code: str = "TW",
        relevance_language: str = "zh-Hant",
        min_minutes: int = 0,
    ) -> List[VideoItem]:
        queries = infer_queries(command)
        self.logger(f"搜尋關鍵字：{' | '.join(queries)}")

        tw_now      = datetime.now(timezone(timedelta(hours=8)))
        tw_midnight = tw_now.replace(hour=0, minute=0, second=0, microsecond=0)
        tw_week_ago = tw_midnight - timedelta(days=6)
        utc_today   = tw_midnight.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        utc_week    = tw_week_ago.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        candidates: Dict[str, dict] = {}
        for q in queries:
            for pub_after, label in [(utc_today, "今日"), (utc_week, "本週")]:
                try:
                    data = self._request(self.SEARCH_URL, {
                        "part": "snippet", "type": "video",
                        "maxResults": 25, "order": "viewCount",
                        "publishedAfter": pub_after, "q": q,
                        "regionCode": region_code,
                        "relevanceLanguage": relevance_language,
                        "safeSearch": "none",
                    })
                    for item in data.get("items", []) or []:
                        vid = safe_str((item.get("id") or {}).get("videoId"))
                        if vid and vid not in candidates:
                            candidates[vid] = item
                    self.logger(f"  [{label}] 「{q}」→ {len(data.get('items', []))} 筆")
                except Exception as e:
                    self.logger(f"  ⚠ 搜尋失敗 [{q}]: {e}")

        self.logger(f"合計候選：{len(candidates)} 支")
        if not candidates:
            raise RuntimeError("YouTube 沒有找到任何影片，請確認 API Key 與網路。")

        # Fetch details
        vid_ids = list(candidates.keys())
        details: Dict[str, dict] = {}
        for i in range(0, len(vid_ids), 50):
            batch = vid_ids[i:i + 50]
            data = self._request(self.VIDEOS_URL, {
                "part": "contentDetails,statistics,snippet",
                "id": ",".join(batch), "maxResults": len(batch),
            })
            for item in data.get("items", []) or []:
                details[safe_str(item.get("id"))] = item

        result: List[VideoItem] = []
        for vid in vid_ids:
            item = details.get(vid)
            if not item:
                continue
            snippet  = item.get("snippet") or {}
            stats    = item.get("statistics") or {}
            content  = item.get("contentDetails") or {}
            dur      = iso8601_duration_to_seconds(safe_str(content.get("duration")))
            if dur <= 60:
                continue    # skip Shorts
            if min_minutes > 0 and dur < min_minutes * 60:
                continue
            views = parse_views_to_int(stats.get("viewCount"))
            result.append(VideoItem(
                rank=0,
                title=safe_str(snippet.get("title")),
                channel=safe_str(snippet.get("channelTitle")),
                url=f"https://www.youtube.com/watch?v={vid}",
                published_at=safe_str(snippet.get("publishedAt")),
                published_local=to_taipei_display(safe_str(snippet.get("publishedAt"))),
                views_text=format_views(views),
                views_num=views,
                duration_text=seconds_to_hms(dur),
                duration_seconds=dur,
                description=safe_str(snippet.get("description")),
            ))

        result.sort(key=lambda x: x.views_num, reverse=True)
        picked = result[:top_n]
        for idx, v in enumerate(picked, start=1):
            v.rank = idx
        self.logger(f"✅ 最終選取：{len(picked)} 支（依播放量排序）")
        return picked


# =============================================================================
# Investment price fetcher
# =============================================================================
_TW_CODES = {"2330","2317","2454","2308","2382","0050","0056","00878","006208","00919","00929"}
_US_TICKERS = {"AAPL","MSFT","NVDA","GOOGL","META","AMZN","TSLA","AMD","INTC","TSM","SPY","QQQ"}
_ALIAS: Dict[str, str] = {
    "黃金": "GC=F", "金": "GC=F", "原油": "CL=F", "白銀": "SI=F",
    "台積電": "2330", "元大台灣50": "0050", "國泰高股息": "00878",
    "輝達": "NVDA", "蘋果": "AAPL", "特斯拉": "TSLA",
    "台灣加權": "^TWII", "加權指數": "^TWII",
    "S&P500": "^GSPC", "NASDAQ": "^IXIC", "道瓊": "^DJI", "費半": "^SOX",
}


def normalize_ticker(raw: str) -> Optional[str]:
    raw = raw.strip().strip("`*_").upper()
    if not raw or raw in {"N/A", "—", "-"}:
        return None
    for k, v in _ALIAS.items():
        if k.upper() == raw:
            return f"{v}.TW" if re.match(r"^\d{4,6}$", v) else v
    if re.match(r"^\d{4,6}$", raw):
        return f"{raw}.TW"
    if raw.endswith(".TW") or raw.startswith("^") or "=" in raw:
        return raw
    if raw in _US_TICKERS:
        return raw
    return None


def fetch_prices(tickers: List[str]) -> List[dict]:
    if not _HAS_YFINANCE:
        return []
    results = []
    for sym in tickers:
        try:
            t  = yf.Ticker(sym)
            fi = t.fast_info
            price = fi.last_price
            prev  = fi.previous_close
            cur   = fi.currency or ("TWD" if sym.endswith(".TW") else "USD")
            if price is None or price != price:
                raise ValueError("NaN")
            chg = price - (prev or price)
            pct = (chg / prev * 100) if prev else 0.0
            try:
                name = t.info.get("shortName") or t.info.get("longName") or sym
            except Exception:
                name = sym
            results.append({
                "symbol": sym, "name": name, "price": price,
                "change": chg, "change_pct": pct, "currency": cur,
            })
            log.info(f"  報價 {sym}: {price:.2f} {cur} ({pct:+.2f}%)")
        except Exception as e:
            log.warning(f"  報價失敗 {sym}: {e}")
            results.append({"symbol": sym, "error": str(e)[:60]})
    return results


# =============================================================================
# HTML Report Generator
# =============================================================================
def build_full_html_report(
    analysis_date: str,
    command: str,
    videos: List[VideoItem],
    summary_md: str,
    report_md: str,
    price_results: List[dict],
    infographic_path: Optional[Path] = None,
) -> str:
    now_tw = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M:%S")

    def _md_to_html(md: str) -> str:
        """Very lightweight Markdown → HTML (headings, bold, tables, lists)."""
        if not md:
            return "<p>（無內容）</p>"
        lines = md.splitlines()
        html_lines = []
        in_table = False
        in_list  = False
        for line in lines:
            # Table
            if "|" in line and line.strip().startswith("|"):
                if not in_table:
                    html_lines.append("<table>")
                    in_table = True
                cells = [c.strip() for c in line.strip().strip("|").split("|")]
                if all(re.match(r"^[-:]+$", c) for c in cells if c):
                    continue
                tag = "th" if not any("<td>" in l for l in html_lines[-5:]) else "td"
                html_lines.append(
                    "<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</tr>"
                )
                continue
            if in_table:
                html_lines.append("</table>")
                in_table = False

            # Close list
            if in_list and not line.strip().startswith("- ") and not line.strip().startswith("* "):
                html_lines.append("</ul>")
                in_list = False

            # Headings
            m = re.match(r"^(#{1,4})\s+(.+)$", line)
            if m:
                lvl = len(m.group(1)) + 1   # h2–h5
                html_lines.append(f"<h{lvl}>{m.group(2)}</h{lvl}>")
                continue

            # List item
            m = re.match(r"^[-*]\s+(.+)$", line)
            if m:
                if not in_list:
                    html_lines.append("<ul>")
                    in_list = True
                html_lines.append(f"<li>{m.group(1)}</li>")
                continue

            # Bold / inline
            processed = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", line)
            processed = re.sub(r"`(.+?)`", r"<code>\1</code>", processed)
            if processed.strip():
                html_lines.append(f"<p>{processed}</p>")

        if in_table:
            html_lines.append("</table>")
        if in_list:
            html_lines.append("</ul>")
        return "\n".join(html_lines)

    # ── Price cards ──────────────────────────────────────────────────────────
    def _price_cards() -> str:
        if not price_results:
            return "<p style='color:#9CA3AF'>未查詢報價（未設定 yfinance 或無標的）</p>"
        cards = []
        for r in price_results:
            if "error" in r:
                cards.append(
                    f'<div class="pcard err">'
                    f'<div class="psym">{r["symbol"]}</div>'
                    f'<div class="perr">❌ {r["error"]}</div>'
                    f'</div>'
                )
                continue
            chg = r["change"]
            pct = r["change_pct"]
            cls = "up" if chg > 0 else ("down" if chg < 0 else "flat")
            arrow = "▲" if chg >= 0 else "▼"
            sign  = "+" if chg >= 0 else ""
            price = r["price"]
            p_str = f"{price:,.2f}" if price >= 1000 else (f"{price:.2f}" if price >= 10 else f"{price:.4f}")
            bar_w = min(100, max(5, abs(pct) * 5))
            cards.append(
                f'<div class="pcard {cls}">'
                f'<div class="psym">{r["symbol"]}</div>'
                f'<div class="pname">{r.get("name","")[:18]}</div>'
                f'<div class="pprice">{p_str}</div>'
                f'<div class="pchg">{arrow} {sign}{chg:.2f} &nbsp; {sign}{pct:.2f}%</div>'
                f'<div class="pbar"><div class="pbari" style="width:{bar_w:.0f}%"></div></div>'
                f'<div class="pcur">{r.get("currency","")}</div>'
                f'</div>'
            )
        return "\n".join(cards)

    # ── Video rows ────────────────────────────────────────────────────────────
    def _video_rows() -> str:
        rows = []
        for v in videos:
            rows.append(
                f'<tr>'
                f'<td class="rank">{v.rank}</td>'
                f'<td><a href="{v.url}" target="_blank">{v.title}</a></td>'
                f'<td>{v.channel}</td>'
                f'<td>{v.published_local}</td>'
                f'<td class="num">{v.views_text}</td>'
                f'<td>{v.duration_text}</td>'
                f'</tr>'
            )
        return "\n".join(rows)

    summary_html = _md_to_html(summary_md)
    report_html  = _md_to_html(report_md)
    price_html   = _price_cards()
    video_html   = _video_rows()

    # ── Infographic: embed as base64 so it shows in email ─────────────────────
    infographic_section = ""
    if infographic_path and Path(infographic_path).exists():
        try:
            img_bytes = Path(infographic_path).read_bytes()
            img_b64   = base64.b64encode(img_bytes).decode()
            infographic_section = (
                '<div class="section">'
                '<h2>🖼️ Infographic（NotebookLM 資訊圖表）</h2>'
                '<div style="background:var(--bg2);border:1px solid var(--border);'
                'border-radius:8px;padding:16px;text-align:center;">'
                f'<img src="data:image/png;base64,{img_b64}" '
                'style="max-width:100%;height:auto;border-radius:6px;" '
                'alt="NotebookLM Infographic">'
                '</div></div>'
            )
        except Exception as _e:
            infographic_section = (
                f'<div class="section"><h2>🖼️ Infographic</h2>'
                f'<p style="color:var(--red)">Infographic 嵌入失敗：{_e}</p></div>'
            )
    else:
        infographic_section = (
            '<div class="section"><h2>🖼️ Infographic（NotebookLM 資訊圖表）</h2>'
            '<p style="color:var(--sub)">本次未產生 Infographic</p></div>'
        )

    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>投資分析日報 – {analysis_date}</title>
<style>
  :root {{
    --bg:#0d1117; --bg2:#161b22; --bg3:#21262d;
    --border:#30363d; --text:#e6edf3; --sub:#8b949e;
    --green:#3fb950; --red:#f85149; --gold:#d29922;
    --blue:#58a6ff; --purple:#bc8cff;
  }}
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:"Microsoft JhengHei UI","Segoe UI",sans-serif;
          background:var(--bg); color:var(--text); font-size:14px; }}
  .page {{ max-width:1200px; margin:0 auto; padding:20px; }}

  /* Header */
  .header {{ border-bottom:2px solid var(--blue); padding-bottom:14px; margin-bottom:20px; }}
  .header h1 {{ font-size:22px; color:var(--blue); }}
  .header .meta {{ font-size:11px; color:var(--sub); margin-top:4px; }}
  .header .cmd {{ background:var(--bg2); border:1px solid var(--border);
                   border-radius:6px; padding:8px 12px; margin-top:10px;
                   font-size:12px; color:var(--gold); }}

  /* Section */
  .section {{ margin-bottom:32px; }}
  .section h2 {{ font-size:16px; color:var(--blue); border-left:3px solid var(--blue);
                  padding-left:10px; margin-bottom:14px; }}

  /* Summary / Report content */
  .content {{ background:var(--bg2); border:1px solid var(--border);
               border-radius:8px; padding:20px; line-height:1.7; }}
  .content h2,.content h3,.content h4 {{ color:var(--gold); margin:14px 0 8px; }}
  .content p {{ margin-bottom:8px; color:var(--text); }}
  .content ul {{ padding-left:20px; margin-bottom:8px; }}
  .content li {{ margin-bottom:4px; }}
  .content strong {{ color:var(--green); }}
  .content code {{ background:var(--bg3); padding:2px 6px; border-radius:4px;
                    font-family:Consolas,monospace; font-size:12px; }}
  .content table {{ width:100%; border-collapse:collapse; margin:12px 0; }}
  .content th {{ background:var(--bg3); color:var(--blue); padding:8px;
                  text-align:left; border:1px solid var(--border); }}
  .content td {{ padding:7px 8px; border:1px solid var(--border);
                  vertical-align:top; }}
  .content tr:nth-child(even) {{ background:var(--bg3); }}

  /* Price cards */
  .price-grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(180px,1fr));
                  gap:10px; }}
  .pcard {{ background:var(--bg2); border:1px solid var(--border);
             border-radius:8px; padding:14px; position:relative; overflow:hidden; }}
  .pcard::before {{ content:''; position:absolute; top:0; left:0; right:0; height:3px; }}
  .pcard.up::before   {{ background:var(--green); }}
  .pcard.down::before {{ background:var(--red); }}
  .pcard.flat::before {{ background:var(--border); }}
  .psym  {{ font-size:15px; font-weight:800; }}
  .pname {{ font-size:10px; color:var(--sub); margin:2px 0 8px;
             white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
  .pprice {{ font-size:20px; font-weight:700; }}
  .pcard.up   .pprice {{ color:var(--green); }}
  .pcard.down .pprice {{ color:var(--red); }}
  .pchg  {{ font-size:12px; margin:4px 0; }}
  .pcard.up   .pchg {{ color:var(--green); }}
  .pcard.down .pchg {{ color:var(--red); }}
  .pbar  {{ background:var(--bg3); height:3px; border-radius:2px; margin:6px 0; }}
  .pbari {{ height:100%; border-radius:2px; }}
  .pcard.up   .pbari {{ background:var(--green); }}
  .pcard.down .pbari {{ background:var(--red); }}
  .pcur  {{ font-size:10px; color:var(--sub); }}
  .perr  {{ font-size:11px; color:var(--red); }}
  .pcard.err {{ opacity:.55; }}

  /* Video table */
  .vtable {{ width:100%; border-collapse:collapse; }}
  .vtable th {{ background:var(--bg3); color:var(--blue); padding:9px 10px;
                 text-align:left; border:1px solid var(--border); font-size:13px; }}
  .vtable td {{ padding:8px 10px; border:1px solid var(--border);
                 vertical-align:top; font-size:13px; }}
  .vtable tr:nth-child(even) {{ background:var(--bg2); }}
  .vtable a {{ color:var(--blue); text-decoration:none; }}
  .vtable a:hover {{ text-decoration:underline; }}
  .vtable .rank {{ text-align:center; font-weight:700; color:var(--gold); width:40px; }}
  .vtable .num  {{ text-align:right; font-family:monospace; }}

  /* Disclaimer */
  .disclaimer {{ background:#1c1a12; border:1px solid var(--gold);
                  border-radius:6px; padding:10px 16px;
                  font-size:11px; color:var(--gold); margin-top:24px; }}
  @media print {{
    body {{ background:#fff; color:#000; }}
    .pcard {{ border:1px solid #ccc; }}
  }}
</style>
</head>
<body>
<div class="page">

  <div class="header">
    <h1>📊 投資分析日報</h1>
    <div class="meta">分析日期：{analysis_date} &nbsp;|&nbsp; 生成時間：{now_tw}（台北時間）&nbsp;|&nbsp; 影片數：{len(videos)} 支</div>
    <div class="cmd">📋 分析指令：{command[:200]}</div>
  </div>

  <!-- 投資報價 -->
  <div class="section">
    <h2>💹 投資標的即時 / 收盤報價</h2>
    <div class="price-grid">
      {price_html}
    </div>
  </div>

  <!-- 影片清單 -->
  <div class="section">
    <h2>🎬 本次分析影片清單（依播放量排序）</h2>
    <div class="content" style="padding:0;overflow:auto">
      <table class="vtable">
        <tr>
          <th class="rank">#</th>
          <th>影片標題</th>
          <th>頻道</th>
          <th>發布時間</th>
          <th>播放數</th>
          <th>長度</th>
        </tr>
        {video_html}
      </table>
    </div>
  </div>

  <!-- 摘要 -->
  <div class="section">
    <h2>📝 分析摘要（NotebookLM Summary）</h2>
    <div class="content">
      {summary_html}
    </div>
  </div>

  <!-- Infographic -->
  {infographic_section}

  <!-- Report -->
  <div class="section">
    <h2>📄 深度報告（NotebookLM Report）</h2>
    <div class="content">
      {report_html}
    </div>
  </div>

  <div class="disclaimer">
    ⚠️ 本報告由 AI 自動生成，所有資訊僅供參考，不構成投資建議。投資有風險，請自行評估。
  </div>

</div>
</body>
</html>"""


# =============================================================================
# NotebookLM pipeline (async)
# =============================================================================
async def run_notebooklm_pipeline(
    notebook_id: str,
    videos: List[VideoItem],
    command: str,
    logger: Callable[[str], None] = log.info,
) -> dict:
    """Register YouTube sources, generate summary + report."""
    if not _HAS_NOTEBOOKLM:
        logger("⚠ notebooklm 套件未安裝，略過 NotebookLM 步驟")
        return {"summary_md": "（未執行 NotebookLM）", "report_md": "（未執行 NotebookLM）"}

    source_rows = [asdict(v) | {"source_kind": "YouTube", "status_text": "待加入"} for v in videos]
    summary_md = report_md = ""

    async with await NotebookLMClient.from_storage() as client:
        # Register sources
        added = 0
        for row in source_rows:
            url = row.get("url", "")
            try:
                logger(f"  加入來源：{url[:60]}")
                await client.sources.add_url(notebook_id, url, wait=True)
                row["status_text"] = "已加入"
                added += 1
            except Exception as e:
                row["status_text"] = f"失敗：{str(e)[:60]}"
                logger(f"  ⚠ 加入失敗：{url} → {e}")

        logger(f"✅ 來源登錄完成：{added}/{len(source_rows)}")

        # Summary
        logger("產生摘要 Markdown…")
        prompt = (
            f"使用者需求：{command}\n\n"
            "請用繁體中文整理本次分析重點，包含：\n"
            "1. 6–10 點條列重點\n"
            "2. 提到的股票/ETF彙整表（代號、觀點、目標價）\n"
            "3. 博主共識與分歧\n"
            "4. 風險提醒\n"
            "輸出純 Markdown，不要輸出 JSON。"
        )
        result = await client.chat.ask(notebook_id, prompt)
        summary_md = (result.answer or "").strip()

        # Report
        logger("產生 Report…")
        try:
            fmt = getattr(ReportFormat, "BRIEFING_DOC", None) if ReportFormat else None
            extra = "請用繁體中文撰寫，適合主管閱讀。"
            if fmt:
                status = await client.artifacts.generate_report(
                    notebook_id, report_format=fmt, language="zh_Hant",
                    extra_instructions=extra)
            else:
                status = await client.artifacts.generate_report(notebook_id)
            await client.artifacts.wait_for_completion(
                notebook_id, status.task_id, timeout=600, initial_interval=5)
            reports = await client.artifacts.list_reports(notebook_id)
            if reports:
                import tempfile, os
                tmp = Path(tempfile.mktemp(suffix=".md"))
                await client.artifacts.download_report(notebook_id, str(tmp))
                report_md = tmp.read_text(encoding="utf-8", errors="ignore")
                tmp.unlink(missing_ok=True)
        except Exception as e:
            logger(f"  ⚠ Report 產生失敗：{e}")
            report_md = f"Report 產生失敗：{e}"

        # Infographic
        infographic_path: Optional[Path] = None
        logger("產生 Infographic…")
        try:
            orient = getattr(InfographicOrientation, "PORTRAIT", None) if InfographicOrientation else None
            detail = getattr(InfographicDetail, "DETAILED", None) if InfographicDetail else None
            instructions = (
                "請用繁體中文生成資訊圖表。"
                "風格清楚、商業報告感，適合主管閱讀。"
                "重點包含：股票/ETF彙整、博主共識、風險警示。"
            )
            # Try with full params, fall back to minimal
            info_status = None
            for attempt_fn in [
                lambda: client.artifacts.generate_infographic(
                    notebook_id, instructions=instructions,
                    language="zh_Hant", orientation=orient, detail=detail),
                lambda: client.artifacts.generate_infographic(
                    notebook_id, instructions=instructions),
                lambda: client.artifacts.generate_infographic(notebook_id),
            ]:
                try:
                    info_status = await attempt_fn()
                    break
                except TypeError:
                    continue

            if info_status:
                await client.artifacts.wait_for_completion(
                    notebook_id, info_status.task_id, timeout=600, initial_interval=5)
                import tempfile
                tmp_png = Path(tempfile.mktemp(suffix=".png"))
                await client.artifacts.download_infographic(notebook_id, str(tmp_png))
                if tmp_png.exists() and tmp_png.stat().st_size > 0:
                    infographic_path = tmp_png
                    logger(f"✅ Infographic 已下載：{tmp_png.stat().st_size // 1024} KB")
                else:
                    logger("⚠ Infographic 下載後檔案為空")
        except Exception as e:
            logger(f"  ⚠ Infographic 產生失敗：{e}")

    return {
        "summary_md": summary_md,
        "report_md": report_md,
        "infographic_path": infographic_path,
    }


# =============================================================================
# Delivery: Email
# =============================================================================
def send_email(
    to_list: List[str],
    from_addr: str,
    app_password: str,
    subject: str,
    html_body: str,
    attachments: List[Path] = [],
):
    if not to_list or not from_addr or not app_password:
        log.warning("Email 設定不完整，略過寄送")
        return
    try:
        msg = MIMEMultipart("mixed")
        msg["Subject"] = subject
        msg["From"]    = from_addr
        msg["To"]      = ", ".join(to_list)
        msg.attach(MIMEText(html_body, "html", "utf-8"))
        for path in attachments:
            if path.exists():
                with open(path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", "attachment",
                                filename=path.name)
                msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(from_addr, app_password)
            smtp.sendmail(from_addr, to_list, msg.as_bytes())
        log.info(f"✅ Email 已寄出 → {', '.join(to_list)}")
    except Exception as e:
        log.error(f"❌ Email 寄送失敗：{e}")


# =============================================================================
# Delivery: LINE Notify
# =============================================================================
def send_line_notify(token: str, message: str):
    if not token:
        log.warning("LINE_NOTIFY_TOKEN 未設定，略過")
        return
    try:
        resp = requests.post(
            "https://notify-api.line.me/api/notify",
            headers={"Authorization": f"Bearer {token}"},
            data={"message": message},
            timeout=15,
        )
        if resp.status_code == 200:
            log.info("✅ LINE Notify 已推送")
        else:
            log.warning(f"LINE Notify 失敗：{resp.status_code} {resp.text[:100]}")
    except Exception as e:
        log.error(f"LINE Notify 錯誤：{e}")


# =============================================================================
# Delivery: Google Drive
# =============================================================================
def upload_to_gdrive(
    sa_json_b64: str,
    folder_id: str,
    file_path: Path,
    mime_type: str = "text/html",
) -> Optional[str]:
    if not sa_json_b64 or not folder_id:
        log.warning("Google Drive 設定不完整，略過上傳")
        return None
    if not _HAS_GDRIVE:
        log.warning("google-api-python-client 未安裝，略過 Drive 上傳")
        return None
    if not file_path.exists():
        log.warning(f"上傳檔案不存在：{file_path}")
        return None
    try:
        sa_json = json.loads(base64.b64decode(sa_json_b64).decode())
        creds   = service_account.Credentials.from_service_account_info(
            sa_json, scopes=["https://www.googleapis.com/auth/drive.file"])
        service = _gdrive_build("drive", "v3", credentials=creds, cache_discovery=False)
        meta    = {"name": file_path.name, "parents": [folder_id]}
        media   = MediaFileUpload(str(file_path), mimetype=mime_type)
        result  = service.files().create(
            body=meta, media_body=media, fields="id,webViewLink").execute()
        link = result.get("webViewLink", "")
        log.info(f"✅ Google Drive 上傳完成：{link}")
        return link
    except Exception as e:
        log.error(f"❌ Google Drive 上傳失敗：{e}")
        return None


# =============================================================================
# Main entry point
# =============================================================================
def setup_storage_state(b64: str) -> bool:
    """Decode and write storage_state.json from base64 env var."""
    if not b64:
        return False
    try:
        data = base64.b64decode(b64)
        dest = Path.home() / ".notebooklm"
        dest.mkdir(parents=True, exist_ok=True)
        (dest / "storage_state.json").write_bytes(data)
        log.info(f"storage_state.json 已寫入：{dest / 'storage_state.json'}")
        return True
    except Exception as e:
        log.error(f"storage_state.json 寫入失敗：{e}")
        return False


async def main_async(cfg: Config):
    analysis_date = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d")
    output_dir    = Path("outputs") / datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d_%H%M%S")
    output_dir.mkdir(parents=True, exist_ok=True)

    log.info("=" * 60)
    log.info(f"📅 分析日期：{analysis_date}")
    log.info(f"📋 分析指令：{cfg.analysis_command[:80]}")
    log.info(f"📁 輸出目錄：{output_dir}")
    log.info("=" * 60)

    # ── Setup storage state ──────────────────────────────────────────────────
    setup_storage_state(cfg.storage_state_b64)

    # ── YouTube search ───────────────────────────────────────────────────────
    log.info("🔍 Step 1: YouTube 搜尋…")
    videos: List[VideoItem] = []
    try:
        yt      = YouTubeSearcher(cfg.youtube_api_key, log.info)
        videos  = yt.search_videos(
            cfg.analysis_command,
            top_n=cfg.top_n,
            region_code=cfg.region_code,
            relevance_language=cfg.language,
            min_minutes=cfg.min_minutes,
        )
        log.info(f"找到 {len(videos)} 支影片")
    except Exception as e:
        log.error(f"YouTube 搜尋失敗：{e}")

    # ── Investment prices ────────────────────────────────────────────────────
    log.info("💹 Step 2: 查詢投資報價…")
    price_results: List[dict] = []
    if _HAS_YFINANCE:
        # Auto-detect tickers from command
        tickers: List[str] = []
        for kw, sym in {
            "台積電": "2330.TW", "0050": "0050.TW", "0056": "0056.TW",
            "00878": "00878.TW", "NVDA": "NVDA", "AAPL": "AAPL",
            "黃金": "GC=F", "原油": "CL=F",
            "台股": "^TWII", "S&P500": "^GSPC", "美股": "^GSPC",
        }.items():
            if kw.lower() in cfg.analysis_command.lower():
                if sym not in tickers:
                    tickers.append(sym)
        # Add extra tickers from env
        for t in cfg.extra_tickers:
            n = normalize_ticker(t)
            if n and n not in tickers:
                tickers.append(n)
        # Always add index benchmarks
        for idx in ["^TWII", "^GSPC"]:
            if idx not in tickers:
                tickers.append(idx)
        log.info(f"查詢代號：{', '.join(tickers)}")
        price_results = fetch_prices(tickers)

    # ── NotebookLM analysis ──────────────────────────────────────────────────
    log.info("🤖 Step 3: NotebookLM 分析…")
    summary_md = report_md = ""
    infographic_path = None
    if videos and cfg.notebook_id and cfg.storage_state_b64:
        try:
            result     = await run_notebooklm_pipeline(
                cfg.notebook_id, videos, cfg.analysis_command, log.info)
            summary_md       = result.get("summary_md", "")
            report_md        = result.get("report_md", "")
            infographic_path = result.get("infographic_path")
        except Exception as e:
            log.error(f"NotebookLM 分析失敗：{e}\n{traceback.format_exc()}")
            summary_md = f"NotebookLM 分析失敗：{e}"
    else:
        log.warning("略過 NotebookLM（無影片、無 Notebook ID 或無 storage_state）")
        summary_md = "（本次未執行 NotebookLM 分析）"

    # ── Build HTML report ────────────────────────────────────────────────────
    log.info("📄 Step 4: 生成 HTML 報告…")
    html = build_full_html_report(
        analysis_date, cfg.analysis_command,
        videos, summary_md, report_md, price_results,
        infographic_path=infographic_path,
    )
    html_path = output_dir / f"daily_report_{analysis_date.replace('-','')}.html"
    html_path.write_text(html, encoding="utf-8")
    log.info(f"HTML 報告已存：{html_path}")

    # ── Google Drive upload ──────────────────────────────────────────────────
    log.info("☁️  Step 5: 上傳 Google Drive…")
    drive_link = upload_to_gdrive(
        cfg.gdrive_sa_b64, cfg.gdrive_folder_id, html_path)

    # ── LINE Notify ──────────────────────────────────────────────────────────
    log.info("📱 Step 6: LINE Notify…")
    # Build a concise summary for LINE (max 1000 chars)
    first_lines = [l.strip() for l in summary_md.splitlines() if l.strip()][:8]
    line_body = (
        f"\n📊 投資分析日報 {analysis_date}\n"
        f"{'─'*28}\n"
        + "\n".join(first_lines[:6])
        + f"\n{'─'*28}\n"
        f"影片來源：{len(videos)} 支\n"
        + (f"完整報告：{drive_link}" if drive_link else "")
    )
    send_line_notify(cfg.line_token, line_body[:990])

    # ── Email ────────────────────────────────────────────────────────────────
    log.info("📧 Step 7: 寄送 Email…")
    email_subject = f"📊 投資分析日報 {analysis_date}"
    # Build attachments list: HTML always included, PNG if available
    email_attachments = [html_path]
    if infographic_path and Path(infographic_path).exists():
        # Copy infographic to output dir with a clean name
        png_dest = output_dir / f"infographic_{analysis_date.replace('-','')}.png"
        try:
            import shutil
            shutil.copy2(str(infographic_path), str(png_dest))
            email_attachments.append(png_dest)
            log.info(f"Infographic 已加入 Email 附件：{png_dest.name}")
        except Exception as e:
            log.warning(f"Infographic 附件複製失敗：{e}")

    send_email(
        cfg.email_to, cfg.email_from, cfg.email_app_password,
        email_subject, html,
        attachments=email_attachments,
    )

    log.info("=" * 60)
    log.info("🎉 全部完成！")
    log.info(f"   HTML：{html_path}")
    if drive_link:
        log.info(f"   Drive：{drive_link}")
    log.info("=" * 60)
    return str(html_path)


def main():
    cfg    = Config()
    errors = cfg.validate()
    if errors:
        log.warning(f"設定警告（不阻擋執行）：{'; '.join(errors)}")
    asyncio.run(main_async(cfg))


if __name__ == "__main__":
    main()
