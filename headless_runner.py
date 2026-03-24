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
def _load_yaml_config(path: str = "config.yaml") -> dict:
    """Load config.yaml; returns {} if missing or unparsable."""
    try:
        import yaml
        for p in [Path(path), Path(__file__).parent / path]:
            if p.exists():
                log.info(f"設定檔載入：{p}")
                with open(p, encoding="utf-8") as f:
                    return yaml.safe_load(f) or {}
        log.warning(f"找不到設定檔 {path}")
    except Exception as e:
        log.warning(f"設定檔載入失敗：{e}")
    return {}


def _cron_matches_now(cron_expr: str, tolerance_minutes: int = 28) -> bool:
    """
    Returns True if the cron expression matches the current UTC time
    within ±tolerance_minutes.  Supports standard 5-field cron.
    """
    try:
        now = datetime.now(timezone.utc)
        parts = cron_expr.strip().split()
        if len(parts) != 5:
            return False
        c_min, c_hr, c_dom, c_mon, c_dow = parts

        def _match(val, current, min_v, max_v):
            if val == "*":
                return True
            if "-" in val:
                a, b = map(int, val.split("-"))
                return a <= current <= b
            if "," in val:
                return current in map(int, val.split(","))
            if "/" in val:
                base, step = val.split("/")
                step = int(step)
                start = int(base) if base != "*" else min_v
                return (current - start) % step == 0 and current >= start
            return int(val) == current

        # Check with tolerance window
        for delta in range(0, tolerance_minutes + 1):
            t = now - timedelta(minutes=delta)
            if (_match(c_min, t.minute, 0, 59) and
                _match(c_hr,  t.hour,   0, 23) and
                _match(c_dow, t.weekday(), 0, 6)):   # Python: Mon=0, Sun=6
                return True
        return False
    except Exception as e:
        log.warning(f"Cron 解析失敗（{cron_expr}）：{e}")
        return False


def _str_tickers(raw) -> List[str]:
    """Parse ticker list from yaml value; strips comments and quotes."""
    if not raw:
        return []
    result = []
    for item in (raw if isinstance(raw, list) else str(raw).split(",")):
        t = str(item).strip().strip('"\'')
        # strip inline yaml comment
        if " #" in t:
            t = t[:t.index(" #")].strip().strip('"\'')
        if t:
            result.append(t)
    return result


def _str_channels(raw) -> List[str]:
    """Parse channel list from yaml; URL-decodes entries."""
    from urllib.parse import unquote as _uq
    if not raw:
        return []
    result = []
    for item in (raw if isinstance(raw, list) else str(raw).split(",")):
        ch = _uq(str(item).strip().strip('"\''))
        if ch:
            result.append(ch)
    return result


class Config:
    """
    Holds settings for ONE job run.
    Built from config.yaml (jobs[] entry) + GitHub Secrets.
    """
    def __init__(self, job_dict: dict, global_cfg: dict):
        defs = global_cfg.get("defaults", {}) or {}
        _nb  = global_cfg.get("notebooklm", {}) or {}
        _rp  = global_cfg.get("report", {}) or {}
        _yt  = job_dict.get("youtube", {}) or {}
        _pr  = job_dict.get("prices", {}) or {}

        # ── Secrets (always from env) ─────────────────────────────────────────
        self.youtube_api_key    = os.environ.get("YOUTUBE_API_KEY", "").strip()
        self.notebook_id        = os.environ.get("NOTEBOOK_ID", "").strip()
        self.storage_state_b64  = os.environ.get("STORAGE_STATE_B64", "").strip()
        self.email_from         = os.environ.get("EMAIL_FROM", "").strip()
        self.email_app_password = os.environ.get("EMAIL_APP_PASSWORD", "").strip()
        self.line_token         = os.environ.get("LINE_NOTIFY_TOKEN", "").strip()
        self.gdrive_sa_b64      = os.environ.get("GDRIVE_SA_JSON_B64", "").strip()
        self.gdrive_folder_id   = os.environ.get("GDRIVE_FOLDER_ID", "").strip()

        _email_secret = os.environ.get("EMAIL_TO", "").strip()
        self.email_to = [e.strip() for e in _email_secret.split(",") if e.strip()] if _email_secret else []

        # ── Job identity ──────────────────────────────────────────────────────
        self.job_id   = str(job_dict.get("id", "default"))
        self.job_name = str(job_dict.get("name", self.job_id))

        # ── Analysis command: job > env > fallback ────────────────────────────
        _cmd_job    = str(job_dict.get("analysis_command", "") or "").strip()
        _cmd_secret = os.environ.get("ANALYSIS_COMMAND", "").strip()
        self.analysis_command = (
            _cmd_secret or _cmd_job or
            "台股 美股 ETF 投資 財經｜請整理今天最熱門的投資影片。"
        )

        # ── YouTube settings: job > defaults > hardcoded ──────────────────────
        self.top_n       = int(_yt.get("top_n") or defs.get("top_n", 10) or 10)
        self.region_code = str(_yt.get("region_code") or defs.get("region_code", "TW") or "TW")
        self.language    = str(_yt.get("language") or defs.get("language", "zh-Hant") or "zh-Hant")
        self.min_minutes = int(_yt.get("min_minutes") if _yt.get("min_minutes") is not None
                               else defs.get("min_minutes", 0) or 0)
        self.search_hours_back = int(
            job_dict.get("search_hours_back") or defs.get("search_hours_back", 21) or 21)
        self.channel_videos_per = int(
            _yt.get("channel_videos_per") or defs.get("channel_videos_per", 3) or 3)

        # Fixed channels – URL-decoded
        self.fixed_channels: List[str] = _str_channels(_yt.get("fixed_channels", []))

        # Extra tickers – comments/quotes stripped + EXTRA_TICKERS secret
        _tickers_yaml   = _str_tickers(_pr.get("extra_tickers", []))
        _tickers_secret = _str_tickers(os.environ.get("EXTRA_TICKERS", ""))
        seen_t: set = set()
        self.extra_tickers: List[str] = []
        for t in _tickers_yaml + _tickers_secret:
            if t and t not in seen_t:
                seen_t.add(t)
                self.extra_tickers.append(t)

        # NotebookLM
        self.infographic_instructions = str(
            _nb.get("infographic_instructions", "請用繁體中文生成資訊圖表，風格清楚商業感。")
        ).strip()

        # Report
        self.subject_prefix = str(_rp.get("subject_prefix", "📊 投資分析日報")).strip()

        log.info(f"[{self.job_id}] 任務：{self.job_name}")
        log.info(f"  指令（前60字）：{self.analysis_command[:60]}…")
        log.info(f"  時間窗口：{self.search_hours_back} 小時 | Top N：{self.top_n}")
        log.info(f"  固定頻道：{len(self.fixed_channels)} 個")
        log.info(f"  報價代號：{', '.join(self.extra_tickers) or '（自動偵測）'}")

    def validate(self) -> List[str]:
        errors = []
        if not self.youtube_api_key:
            errors.append("YOUTUBE_API_KEY 未設定")
        if not self.notebook_id and _HAS_NOTEBOOKLM:
            errors.append("NOTEBOOK_ID 未設定")
        if not self.storage_state_b64 and _HAS_NOTEBOOKLM:
            errors.append("STORAGE_STATE_B64 未設定")
        return errors


def load_jobs_to_run(job_id_filter: str = "") -> List[tuple]:
    """
    Returns list of (Config, global_cfg) for jobs that should run now.
    job_id_filter: if set, only run that specific job (ignores schedule).
    """
    global_cfg = _load_yaml_config()
    jobs = global_cfg.get("jobs", []) or []

    # Fall back to legacy single-job format
    if not jobs:
        log.warning("config.yaml 沒有 jobs 清單，使用舊版單一任務格式")
        legacy_job = {
            "id": "default",
            "name": "預設任務",
            "enabled": True,
            "analysis_command": global_cfg.get("analysis_command", ""),
            "youtube": global_cfg.get("youtube", {}),
            "prices": global_cfg.get("prices", {}),
            "search_hours_back": (global_cfg.get("youtube", {}) or {}).get("search_hours_back", 21),
        }
        jobs = [legacy_job]

    result = []
    for job in jobs:
        jid     = str(job.get("id", ""))
        enabled = job.get("enabled", True)
        cron    = str(job.get("schedule", "") or "")

        # Filter by job_id if specified
        if job_id_filter and jid != job_id_filter:
            continue

        if not enabled and not job_id_filter:
            log.info(f"  跳過停用任務：{jid}")
            continue

        # Check schedule (only when not manually specified)
        if cron and not job_id_filter:
            if not _cron_matches_now(cron):
                log.info(f"  [{jid}] 尚未到執行時間（cron={cron}），跳過")
                continue

        try:
            cfg = Config(job, global_cfg)
            errs = cfg.validate()
            if errs:
                log.warning(f"  [{jid}] 設定警告：{'; '.join(errs)}")
            result.append(cfg)
        except Exception as e:
            log.error(f"  [{jid}] Config 建立失敗：{e}")

    if not result:
        if job_id_filter:
            log.error(f"找不到 job_id={job_id_filter}")
        else:
            log.info("目前時間沒有需要執行的任務")
    return result


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
    source_type: str = "search"   # "search" | "fixed_channel"


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
        hours_back: int = 21,
    ) -> List[VideoItem]:
        queries = infer_queries(command)
        self.logger(f"搜尋關鍵字：{' | '.join(queries)}")

        tw_now     = datetime.now(timezone(timedelta(hours=8)))
        # Window start: hours_back hours ago (default = yesterday 12:00 if run at 09:00)
        window_start = tw_now - timedelta(hours=hours_back)
        utc_start    = window_start.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.logger(
            f"搜尋時間窗口：{window_start.strftime('%Y-%m-%d %H:%M')} "
            f"~ {tw_now.strftime('%Y-%m-%d %H:%M')}（台北時間）"
        )

        candidates: Dict[str, dict] = {}
        for q in queries:
            for pub_after, label in [(utc_start, f"近{hours_back}小時")]:
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


    def fetch_channel_recent_videos(
        self,
        channel_ref: str,
        hours_back: int = 21,
        max_videos: int = 3,
    ) -> List[VideoItem]:
        """Fetch recent videos from a specific channel within the time window."""
        from urllib.parse import unquote, urlparse as _urlparse

        from urllib.parse import unquote as _url_unquote
        ref = _url_unquote(channel_ref.strip())   # decode %E6%97%A9... → 早見說財經
        if not ref:
            return []

        channel_id = ""
        channel_name = ref

        # Try to get channel ID from URL or @handle
        if ref.startswith("http"):
            try:
                parsed = _urlparse(ref)
                path   = unquote(parsed.path or "")
                m = re.search(r"/channel/(UC[\w-]{20,})", path)
                if m:
                    channel_id = m.group(1)
                if not channel_id:
                    resp = self.session.get(ref, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
                    if resp.status_code == 200:
                        m2 = re.search(r'"channelId":"(UC[\w-]{20,})"', resp.text)
                        if not m2:
                            m2 = re.search(r'"externalId":"(UC[\w-]{20,})"', resp.text)
                        if m2:
                            channel_id = m2.group(1)
                        t = re.search(r'<meta[^>]+property="og:title"[^>]+content="([^"]+)"', resp.text, re.I)
                        if t:
                            channel_name = t.group(1).strip()
            except Exception as e:
                self.logger(f"  解析頻道網址失敗 {ref}: {e}")

        if not channel_id:
            # Search by handle/name
            q = ref.lstrip("@").split("/")[-1]
            try:
                data  = self._request(self.SEARCH_URL, {
                    "part": "snippet", "type": "channel",
                    "maxResults": 3, "q": q, "safeSearch": "none",
                })
                items = data.get("items", []) or []
                if items:
                    best = items[0]
                    channel_id   = safe_str((best.get("snippet") or {}).get("channelId")
                                            or (best.get("id") or {}).get("channelId"))
                    channel_name = safe_str((best.get("snippet") or {}).get("channelTitle") or q)
            except Exception as e:
                self.logger(f"  搜尋頻道失敗 {ref}: {e}")
                return []

        if not channel_id:
            self.logger(f"  ⚠ 找不到頻道 ID：{ref}")
            return []

        self.logger(f"  頻道：{channel_name} ({channel_id})")

        tw_now       = datetime.now(timezone(timedelta(hours=8)))
        window_start = tw_now - timedelta(hours=hours_back)
        utc_start    = window_start.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        try:
            data  = self._request(self.SEARCH_URL, {
                "part": "snippet", "type": "video",
                "channelId": channel_id, "maxResults": 10,
                "order": "date", "publishedAfter": utc_start,
                "safeSearch": "none",
            })
        except Exception as e:
            self.logger(f"  ⚠ 搜尋頻道影片失敗：{e}")
            return []

        vid_ids = [safe_str((item.get("id") or {}).get("videoId"))
                   for item in (data.get("items") or [])
                   if safe_str((item.get("id") or {}).get("videoId"))]

        if not vid_ids:
            self.logger(f"  {channel_name}：此時間窗口內無新影片")
            return []

        # Fetch details
        details: Dict[str, dict] = {}
        det = self._request(self.VIDEOS_URL, {
            "part": "contentDetails,statistics,snippet",
            "id": ",".join(vid_ids[:10]), "maxResults": 10,
        })
        for item in det.get("items", []) or []:
            details[safe_str(item.get("id"))] = item

        result: List[VideoItem] = []
        for vid in vid_ids:
            item = details.get(vid)
            if not item:
                continue
            snippet  = item.get("snippet")        or {}
            stats    = item.get("statistics")     or {}
            content  = item.get("contentDetails") or {}
            dur      = iso8601_duration_to_seconds(safe_str(content.get("duration")))
            if dur <= 60:
                continue   # skip Shorts
            views = parse_views_to_int(stats.get("viewCount"))
            result.append(VideoItem(
                rank=0,
                title=safe_str(snippet.get("title")),
                channel=channel_name,
                url=f"https://www.youtube.com/watch?v={vid}",
                published_at=safe_str(snippet.get("publishedAt")),
                published_local=to_taipei_display(safe_str(snippet.get("publishedAt"))),
                views_text=format_views(views),
                views_num=views,
                duration_text=seconds_to_hms(dur),
                duration_seconds=dur,
                description=safe_str(snippet.get("description")),
                source_type="fixed_channel",
            ))

        result = result[:max_videos]
        self.logger(f"  {channel_name}：取得 {len(result)} 支影片")
        return result


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
            src_badge = ('<span style="color:#f59e0b;font-size:10px;margin-left:4px" '
                         'title="固定頻道來源">📺</span>'
                         if getattr(v, "source_type", "") == "fixed_channel" else "")
            rows.append(
                f'<tr>'
                f'<td class="rank">{v.rank}</td>'
                f'<td><a href="{v.url}" target="_blank">{v.title}</a>{src_badge}</td>'
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
            import tempfile, traceback as _tb
            orient = getattr(InfographicOrientation, "PORTRAIT", None) if InfographicOrientation else None
            detail = getattr(InfographicDetail, "DETAILED", None) if InfographicDetail else None
            instructions = cfg.infographic_instructions if hasattr(cfg, "infographic_instructions") else (
                "請用繁體中文生成資訊圖表，風格清楚商業感。"
            )

            # Snapshot before – to detect the new artifact after generation
            infos_before = []
            try:
                infos_before = await client.artifacts.list_infographics(notebook_id)
                logger(f"  現有 Infographic 數量：{len(infos_before)}")
            except Exception as _e:
                logger(f"  list_infographics (before) 失敗（不影響）：{_e}")

            # Try with progressively simpler params
            info_status = None
            attempts = [
                ("完整參數", lambda: client.artifacts.generate_infographic(
                    notebook_id, instructions=instructions,
                    language="zh_Hant", orientation=orient, detail=detail)),
                ("移除 detail", lambda: client.artifacts.generate_infographic(
                    notebook_id, instructions=instructions, language="zh_Hant")),
                ("instructions only", lambda: client.artifacts.generate_infographic(
                    notebook_id, instructions=instructions)),
                ("bare", lambda: client.artifacts.generate_infographic(notebook_id)),
            ]
            for label, fn in attempts:
                try:
                    logger(f"  嘗試 generate_infographic（{label}）…")
                    info_status = await fn()
                    logger(f"  generate_infographic 回傳 task_id={getattr(info_status, 'task_id', '?')}")
                    break
                except TypeError as te:
                    logger(f"  TypeError，嘗試下一組參數：{te}")
                    continue
                except Exception as ge:
                    logger(f"  generate_infographic 失敗（{label}）：{ge}")
                    break

            if info_status and getattr(info_status, "task_id", None):
                logger(f"  等待 Infographic 完成（task_id={info_status.task_id}）…")
                await client.artifacts.wait_for_completion(
                    notebook_id, info_status.task_id, timeout=600, initial_interval=5)
                logger("  Infographic 已完成，準備下載…")

                # Detect new artifact
                artifact_id = None
                try:
                    infos_after = await client.artifacts.list_infographics(notebook_id)
                    logger(f"  完成後 Infographic 數量：{len(infos_after)}")
                    before_ids = {getattr(i, "id", None) or (i[0] if isinstance(i, (list,tuple)) else None)
                                  for i in infos_before}
                    for item in infos_after:
                        aid = getattr(item, "id", None) or (item[0] if isinstance(item, (list,tuple)) else None)
                        if aid and aid not in before_ids:
                            artifact_id = aid
                            break
                    if not artifact_id and infos_after:
                        item = infos_after[0]
                        artifact_id = getattr(item, "id", None) or (item[0] if isinstance(item, (list,tuple)) else None)
                    logger(f"  Infographic artifact_id={artifact_id}")
                except Exception as _e:
                    logger(f"  list_infographics (after) 失敗：{_e}")

                tmp_png = Path(tempfile.mktemp(suffix=".png"))
                try:
                    if artifact_id:
                        await client.artifacts.download_infographic(
                            notebook_id, str(tmp_png), artifact_id=artifact_id)
                    else:
                        await client.artifacts.download_infographic(notebook_id, str(tmp_png))
                    if tmp_png.exists() and tmp_png.stat().st_size > 0:
                        infographic_path = tmp_png
                        logger(f"✅ Infographic 已下載：{tmp_png.stat().st_size // 1024} KB")
                    else:
                        logger("⚠ Infographic 下載後檔案為空或不存在")
                except Exception as dl_e:
                    logger(f"  download_infographic 失敗：{dl_e}")
                    logger(_tb.format_exc())
            else:
                logger("⚠ generate_infographic 未回傳有效 task_id，略過下載")
        except Exception as e:
            logger(f"  ⚠ Infographic 流程異常：{e}")
            import traceback as _tb2
            logger(_tb2.format_exc())

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
    ts_str        = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d_%H%M%S")
    output_dir    = Path("outputs") / f"{cfg.job_id}_{ts_str}"
    output_dir.mkdir(parents=True, exist_ok=True)

    log.info("=" * 60)
    log.info(f"🏷  任務：{cfg.job_name}  ({cfg.job_id})")
    log.info(f"📅 分析日期：{analysis_date}")
    log.info(f"📋 分析指令：{cfg.analysis_command[:80]}")
    log.info(f"📁 輸出目錄：{output_dir}")
    log.info("=" * 60)

    # ── Setup storage state ──────────────────────────────────────────────────
    setup_storage_state(cfg.storage_state_b64)

    # ── YouTube search ───────────────────────────────────────────────────────
    log.info("🔍 Step 1: YouTube 搜尋…")
    log.info(f"  搜尋時間窗口：過去 {cfg.search_hours_back} 小時")
    videos: List[VideoItem] = []
    try:
        yt     = YouTubeSearcher(cfg.youtube_api_key, log.info)
        videos = yt.search_videos(
            cfg.analysis_command,
            top_n=cfg.top_n,
            region_code=cfg.region_code,
            relevance_language=cfg.language,
            min_minutes=cfg.min_minutes,
            hours_back=cfg.search_hours_back,
        )
        log.info(f"關鍵字搜尋找到 {len(videos)} 支影片")
    except Exception as e:
        log.error(f"YouTube 搜尋失敗：{e}")

    # ── Fixed YouTuber channels ───────────────────────────────────────────────
    if cfg.fixed_channels:
        log.info(f"📺 Step 1b: 抓取固定 YouTuber 頻道（共 {len(cfg.fixed_channels)} 個）…")
        existing_urls = {v.url for v in videos}
        for ch_ref in cfg.fixed_channels:
            try:
                ch_videos = yt.fetch_channel_recent_videos(
                    ch_ref,
                    hours_back=cfg.search_hours_back,
                    max_videos=cfg.channel_videos_per,
                )
                added = 0
                for cv in ch_videos:
                    if cv.url not in existing_urls:
                        videos.append(cv)
                        existing_urls.add(cv.url)
                        added += 1
                log.info(f"  {ch_ref[:50]}：新增 {added} 支")
            except Exception as e:
                log.warning(f"  ⚠ 頻道抓取失敗 {ch_ref[:50]}: {e}")

        # Re-rank
        for idx, v in enumerate(videos, start=1):
            v.rank = idx
        log.info(f"固定頻道加入後，共 {len(videos)} 支影片")

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
    email_subject = f"{cfg.subject_prefix} {cfg.job_name} {analysis_date}"
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
    import argparse
    parser = argparse.ArgumentParser(description="投資日報自動分析")
    parser.add_argument("--job-id", default=os.environ.get("JOB_ID", ""),
                        help="指定執行的任務 ID（留空=自動依排程判斷）")
    parser.add_argument("--list-jobs", action="store_true",
                        help="列出所有任務設定後退出")
    args = parser.parse_args()

    if args.list_jobs:
        global_cfg = _load_yaml_config()
        jobs = global_cfg.get("jobs", []) or []
        print(f"共 {len(jobs)} 個任務：")
        for j in jobs:
            status = "✅ 啟用" if j.get("enabled", True) else "⛔ 停用"
            print(f"  [{j.get('id','?')}] {j.get('name','')} | {status} | cron={j.get('schedule','?')}")
        return

    job_id_filter = args.job_id.strip()
    if job_id_filter:
        log.info(f"手動指定執行任務：{job_id_filter}")
    else:
        log.info("自動模式：依排程判斷需要執行的任務")

    cfgs = load_jobs_to_run(job_id_filter)
    if not cfgs:
        log.info("本次無需執行任何任務，結束。")
        return

    for cfg in cfgs:
        log.info(f"\n{'='*60}\n開始執行任務：{cfg.job_name}\n{'='*60}")
        asyncio.run(main_async(cfg))
        log.info(f"任務 {cfg.job_id} 完成")


if __name__ == "__main__":
    main()
