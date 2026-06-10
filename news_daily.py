"""
news_daily.py — 無 UI 版，供 GitHub Actions 排程執行
功能：
  1. 抓取前一天 00:00 ~ 當下的市府新聞稿
  2. 比對 Google Sheets 已有 URL，只寫入新增的
  3. 用 Gemini API 產生摘要與 FAQ
  4. 將結果 append 到 Google Sheets
環境變數（GitHub Secrets）：
  GEMINI_API_KEY      — Gemini API 金鑰
  GOOGLE_CREDS_JSON   — credentials.json 的完整內容（JSON 字串）
  SHEET_ID            — 輸出用的 Google Sheets ID
"""

import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import urllib3
import requests
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google import genai
from google.genai import types

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── 常數 ──────────────────────────────────────────────────────────────────────
TW_TZ      = timezone(timedelta(hours=8))
BASE_GOV   = "https://www.gov.taipei"
NEWS_URL   = (f"{BASE_GOV}/News.aspx"
              "?n=F0DDAF49B89E9413&sms=72544237BBE4C5F6&PageSize=100")
HEADERS    = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9",
}
SHEET_COLS = ["來源", "日期", "標題", "發布機關", "連結", "內容", "摘要", "FAQ"]
# ─────────────────────────────────────────────────────────────────────────────


# ── Google Sheets 連線 ────────────────────────────────────────────────────────
def connect_sheet(sheet_id: str) -> gspread.Worksheet:
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]

    # 優先用環境變數（GitHub Actions），其次用本地 credentials.json
    creds_json = os.environ.get("GOOGLE_CREDS_JSON")
    if creds_json:
        info = json.loads(creds_json)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            Path(__file__).parent / "credentials.json", scope)

    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)

    # 確保工作表有標題列
    ws = sh.sheet1
    if ws.row_count == 0 or not ws.row_values(1):
        ws.append_row(SHEET_COLS)
    return ws


# ── 日期轉換 ──────────────────────────────────────────────────────────────────
def roc_to_date(s: str) -> datetime | None:
    m = re.search(r"(\d{2,3})[-年](\d{1,2})[-月](\d{1,2})", s.strip())
    if not m:
        return None
    try:
        return datetime(int(m.group(1)) + 1911, int(m.group(2)),
                        int(m.group(3)), tzinfo=TW_TZ)
    except ValueError:
        return None


# ── 新聞列表抓取 ──────────────────────────────────────────────────────────────
def fetch_list(start_dt: datetime, end_dt: datetime) -> list[dict]:
    rows, page = [], 1
    while True:
        print(f"  列表第 {page} 頁…")
        try:
            resp = requests.get(NEWS_URL + f"&page={page}",
                                headers=HEADERS, timeout=20, verify=False)
            resp.encoding = "utf-8"
        except Exception as e:
            print(f"  [錯誤] {e}"); break

        soup = BeautifulSoup(resp.text, "html.parser")
        trs  = soup.select("tbody tr")
        if not trs:
            break

        found_in_page = 0
        oldest_in_page = None

        for tr in trs:
            date_td  = tr.select_one('td[data-title="發布日期"]')
            title_td = tr.select_one('td[data-title="標題"] a')
            dept_td  = tr.select_one('td[data-title="發布機關"]')
            if not date_td or not title_td:
                continue

            date_str  = date_td.get_text(strip=True)
            pub       = roc_to_date(date_str)
            if not pub:
                continue

            # 記錄這頁最舊的日期，用於判斷是否繼續翻頁
            if oldest_in_page is None or pub < oldest_in_page:
                oldest_in_page = pub

            pub_day   = pub.replace(hour=0, minute=0, second=0)
            start_day = start_dt.replace(hour=0, minute=0, second=0)
            end_day   = end_dt.replace(hour=0, minute=0, second=0)
            if pub_day < start_day or pub_day > end_day:
                continue

            title_text = title_td.get_text(strip=True)
            dept_text  = dept_td.get_text(strip=True) if dept_td else ""
            raw  = title_td.get("href", "")
            href = raw if raw.startswith("http") else f"{BASE_GOV}/{raw.lstrip('/')}"
            found_in_page += 1
            rows.append({"來源": "市府新聞稿", "日期": pub.strftime("%Y-%m-%d"),
                         "標題": title_text, "發布機關": dept_text,
                         "連結": href, "內容": "", "摘要": "", "FAQ": ""})

        # 若這頁最舊的文章已早於起始日，不需再翻頁
        if oldest_in_page and oldest_in_page.replace(hour=0,minute=0,second=0) < start_dt.replace(hour=0,minute=0,second=0):
            break
        if len(trs) < 100:
            break
        page += 1
    return rows


# ── 文章內容抓取 ──────────────────────────────────────────────────────────────
def fetch_content(url: str) -> str:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20, verify=False)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        div  = (soup.find("div", class_=re.compile(
                    r"(news[-_]?content|article|content[-_]?body|rte)", re.I))
                or soup.find("article")
                or soup.find("div", id=re.compile(r"(content|article|main)", re.I)))
        if not div:
            h3  = soup.find("h3")
            div = h3.find_parent("div") if h3 else soup.find("body")
        if div:
            for t in div.find_all(["script", "style", "nav", "header", "footer"]):
                t.decompose()
            paras = div.find_all("p")
            text  = "\n".join(p.get_text(" ", strip=True)
                              for p in paras if p.get_text(strip=True))
            return (text or div.get_text("\n", strip=True))[:5000]
    except Exception as e:
        return f"[抓取失敗：{e}]"
    return ""


# ── LINE 通知 ─────────────────────────────────────────────────────────────────
def send_line(rows: list[dict], start_dt: datetime, end_dt: datetime) -> None:
    token = os.environ.get("LINE_CHANNEL_TOKEN", "")
    if not token:
        print("LINE 通知未設定，略過。")
        return

    period = f"{start_dt.strftime('%m/%d %H:%M')} ~ {end_dt.strftime('%m/%d %H:%M')}"
    header = f"【臺北市府新聞】{period}　共 {len(rows)} 筆\n"

    # 廣播：所有加好友的人都收到，不需要指定 User ID
    def push(text: str):
        resp = requests.post(
            "https://api.line.me/v2/bot/message/broadcast",
            headers={"Authorization": f"Bearer {token}",
                     "Content-Type": "application/json"},
            json={"messages": [{"type": "text", "text": text}]},
            timeout=15,
        )
        if resp.status_code != 200:
            print(f"LINE 通知失敗：{resp.status_code} {resp.text[:100]}")

    try:
        chunk, chunk_rows = header, []
        for r in rows:
            item = (f"📅 {r['日期']}\n"
                    f"📌 {r['標題']}\n"
                    f"🏢 {r['發布機關']}\n"
                    f"🔗 {r['連結']}\n\n")
            if len(chunk) + len(item) > 4800:
                push(chunk.strip())
                chunk = ""
            chunk += item

        if chunk.strip():
            push(chunk.strip())
        print(f"LINE 通知已發送（共 {len(rows)} 筆）")
    except Exception as e:
        print(f"LINE 通知錯誤：{e}")


# ── Gemini AI ─────────────────────────────────────────────────────────────────
def gemini_generate(prompt: str, model_name: str = "gemini-2.0-flash-lite") -> str:
    client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY", ""))
    for attempt in range(3):
        try:
            result = client.models.generate_content(model=model_name, contents=prompt)
            return result.text.strip()
        except Exception as e:
            msg = str(e)
            if "429" in msg and attempt < 2:
                wait = 30 * (attempt + 1)   # 30s, 60s
                print(f"  [429 限流] 等待 {wait} 秒後重試…")
                time.sleep(wait)
            else:
                return f"[AI 失敗：{msg[:120]}]"
    return "[AI 失敗：重試次數用盡]"


def make_summary(content: str, model_name: str) -> str:
    if not content or content.startswith("[") or len(content) < 20:
        return ""
    return gemini_generate(
        "你是政府公文摘要助理，請用繁體中文將以下新聞稿濃縮成3～5句話的摘要，"
        "重點包含：事件主體、核心內容、時間地點（若有）。直接輸出摘要，不要加標題或說明。\n\n"
        f"{content[:3000]}", model_name)


def make_faq(content: str, model_name: str) -> str:
    if not content or content.startswith("[") or len(content) < 20:
        return ""
    return gemini_generate(
        "你是政府新聞稿 FAQ 整理助理。請仔細閱讀以下新聞稿全文，"
        "用繁體中文依照下列固定格式逐項回答。"
        "每一項請直接從原文中找出對應資訊並摘錄，若原文未提及則寫「未提及」。\n\n"
        "Q1: 這則公告的主要內容是什麼？\nA1: （說明事件核心或政策重點，2～4句）\n\n"
        "Q2: 時間是什麼時候？\nA2: （完整摘錄所有日期與時段）\n\n"
        "Q3: 地點在哪裡？\nA3: （完整摘錄所有路名、路口、巷弄、地址，不可省略）\n\n"
        "Q4: 影響範圍或注意事項是什麼？\nA4: （封路範圍、替代路線、管制措施等，逐點列出）\n\n"
        "Q5: 適用對象或申請資格是什麼？\nA5: （對象、身分條件等）\n\n"
        "Q6: 如何報名、申請或參與？\nA6: （報名方式、申請流程、網址等）\n\n"
        "Q7: 費用是多少？有無補助或優惠？\nA7: （收費、補助金額等）\n\n"
        "Q8: 如何聯絡或洽詢？\nA8: （電話、分機、網址、承辦單位，完整列出）\n\n"
        "注意：Q3 務必直接引用原文路名，不可用「施工現場」帶過。全程使用繁體中文。\n"
        "直接輸出以上格式，不要加任何額外說明。\n\n"
        f"===新聞稿內容===\n{content[:3000]}", model_name)


# ── 主程式 ────────────────────────────────────────────────────────────────────
def main():
    # --- 設定 ---
    gemini_key = os.environ.get("GEMINI_API_KEY", "")
    sheet_id   = os.environ.get("SHEET_ID", "")
    if not gemini_key or not sheet_id:
        # 本機開發：從同目錄的 .env 讀取
        env_path = Path(__file__).parent / ".env"
        if env_path.exists():
            for line in env_path.read_text(encoding="utf-8").splitlines():
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())
            gemini_key = os.environ.get("GEMINI_API_KEY", "")
            sheet_id   = os.environ.get("SHEET_ID", "")
    if not gemini_key:
        sys.exit("錯誤：請設定 GEMINI_API_KEY 環境變數")
    if not sheet_id:
        sys.exit("錯誤：請設定 SHEET_ID 環境變數")
    use_ai     = os.environ.get("USE_AI", "true").lower() == "true"
    print(f"AI 功能：{'開啟（模型：' + os.environ.get('GEMINI_MODEL','') + '）' if use_ai else '關閉'}")
    gemini_model = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash")

    now      = datetime.now(tz=TW_TZ)
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    # 06:00 之前的第一班，把昨天也一起抓（避免深夜發稿漏抓）
    if now.hour < 7:
        start_dt = today_start - timedelta(days=1)
    else:
        start_dt = today_start
    end_dt = now
    print(f"抓取期間：{start_dt.strftime('%Y-%m-%d %H:%M')} ~ {end_dt.strftime('%Y-%m-%d %H:%M')}")

    # --- 連接 Sheets，取已有 URL ---
    print("連接 Google Sheets…")
    ws = connect_sheet(sheet_id)

    # --- 刪除 4 天前的舊資料（不含第1列標題，不含4天前當天） ---
    cutoff = (datetime.now(tz=TW_TZ) - timedelta(days=2)).strftime("%Y-%m-%d")
    all_rows = ws.get_all_values()
    rows_to_delete = []
    for i, row in enumerate(all_rows):
        if i == 0:          # 跳過標題列
            continue
        if len(row) >= 2 and row[1] != "" and row[1] < cutoff:
            rows_to_delete.append(i + 1)   # Sheets 列號從 1 開始
    for row_num in sorted(rows_to_delete, reverse=True):   # 從底部往上刪，避免列號偏移
        ws.delete_rows(row_num)
    if rows_to_delete:
        print(f"已刪除 {len(rows_to_delete)} 筆過期資料（早於 {cutoff}）")

    existing_urls = set(ws.col_values(5))   # 第5欄 = 連結

    # --- 抓新聞列表 ---
    print("抓取新聞列表…")
    rows = fetch_list(start_dt, end_dt)
    print(f"列表共 {len(rows)} 筆")

    # --- 過濾已存在的 ---
    new_rows = [r for r in rows if r["連結"] not in existing_urls]
    print(f"排除重複後，新增 {len(new_rows)} 筆")
    if not new_rows:
        print("無新增內容，結束。")
        return

    # --- 抓內容 + AI ---
    total = len(new_rows)
    for i, row in enumerate(new_rows, 1):
        print(f"  [{i}/{total}] 抓內容：{row['標題'][:30]}…")
        row["內容"] = fetch_content(row["連結"])
        time.sleep(0.5)

        if use_ai:
            print(f"  [{i}/{total}] AI 摘要…")
            row["摘要"] = make_summary(row["內容"], gemini_model)
            time.sleep(1)
            print(f"  [{i}/{total}] AI FAQ…")
            row["FAQ"]  = make_faq(row["內容"], gemini_model)
            time.sleep(1)

    # --- 寫入 Sheets ---
    print("寫入 Google Sheets…")
    append_data = [[r[c] for c in SHEET_COLS] for r in new_rows]
    ws.append_rows(append_data, value_input_option="RAW")
    print(f"完成，共寫入 {len(append_data)} 筆")

    # --- LINE 通知 ---
    send_line(new_rows, start_dt, end_dt)


if __name__ == "__main__":
    main()
