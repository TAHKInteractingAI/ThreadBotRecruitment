import os
import sys
import time
import json
import random
import traceback
import asyncio
import functools
import unicodedata
import requests
import regex as re
import gspread
import pytz

from datetime import datetime
from PIL import Image
from google.oauth2.service_account import Credentials
from google.auth.exceptions import TransportError
from gspread.utils import rowcol_to_a1
from playwright.async_api import async_playwright

# Windows async event loop fix
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# =========================
# CONFIGURATION
# =========================
BASE_DIR = os.getcwd()
CREDENTIAL_FILE = "credentials.json"
RECRUIT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1o1LRTqzxje8XUBh_oulx5lSnFuId_QLSdAT-82IENGI/edit?usp=sharing"

SHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_MAX_RETRIES = 5
SHEET_RETRY_BASE_DELAY = 3
SHEET_HTTP_TIMEOUT = 30

RECRUIT_TAB_NAME = "Recruitment"
ACCOUNT_TAB_NAME = "Accounts"

COL_POSITION = "Position"
COL_JOB_CONTENT = "Job Content"
COL_THREAD_CONTENT = "Thread Content"
COL_TOPIC = "Topic"
COL_IMAGE = "Image URL"
COL_POSTED = "Posted"
COL_LINK_POST = "Link post"
COL_DATE = "Date"
COL_ACCOUNTS_CODE = "AccountsCode"

COL_ACC_USERNAME = "Username"

COL_ACC_ACTIVE = "Active"
INACTIVE_VALUES = {"NO", "N", "FALSE", "0", "OFF", "INACTIVE", "DISABLED"}

MAX_POSTS_PER_RUN = 999
THREADS_URL = "https://www.threads.com"

THREADS_CHAR_LIMIT = 480

MIN_TAIL_CHARS = 120

MERGE_OVERFLOW_WITH_THREAD = True

STATUS_YES = "YES"
STATUS_ERROR = "ERROR"
STATUS_REPLY_ERROR = "REPLY_ERROR"
STATUS_CHECK = "CHECK_MANUAL"
YES_WITHOUT_LINK_ACTION = "retry"
STATUS_YES_NO_LINK = "YES_NO_LINK"
STATUS_ACCOUNT_DISABLED = "Account disabled"

HEADLESS = os.environ.get("HEADLESS", "1") != "0"

REPLY_MAX_ATTEMPTS = 3

VERIFY_TYPED_TEXT = True
TYPE_MAX_ATTEMPTS = 3

NAV_TIMEOUT_MS = 60000
DEFAULT_TIMEOUT_MS = 30000
GOTO_ATTEMPTS = 3

DISABLED_CACHE_FILE = os.path.join(BASE_DIR, "disabled_accounts.json")
DISABLED_RECHECK_HOURS = 24

UNSYNCED_FILE = os.path.join(BASE_DIR, "unsynced_status.jsonl")

DISABLED_MARKERS = [
    "đã vô hiệu hóa tài khoản",
    "vô hiệu hóa tài khoản của bạn",
    "không có quyền truy cập vào",
    "we disabled your account",
    "we've disabled your account",
    "your account has been disabled",
    "account has been suspended",
    "tài khoản của bạn đã bị đình chỉ",
]


class AccountDisabledError(Exception):
    """Tài khoản Threads bị vô hiệu hóa/khóa, không thể dùng được nữa."""


class PostUnverifiedError(Exception):
    """Đã bấm đăng nhưng không xác minh được bài có lên hay chưa."""


POST_DELAY_RANGE = (3, 6)
AFTER_POST_DELAY = (5, 8)
DELAY_BETWEEN_POSTS = (60, 180)


# =========================
# ULTILS
# =========================
def normalize_threads_content(text: str) -> str:
    if not text:
        return ""
    t = unicodedata.normalize("NFC", text).strip()
    t = t.replace("\r\n", "\n")
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def split_text(text: str, limit: int) -> list:
    """
    Cắt text thành các đoạn <= limit ký tự.
    Ưu tiên cắt ở: hết đoạn (\\n\\n) > xuống dòng (\\n) > hết câu > khoảng trắng > cắt cứng.
    """
    text = (text or "").strip()
    if not text:
        return []
    if len(text) <= limit:
        return [text]

    parts = []
    min_cut = int(limit * 0.4)
    while len(text) > limit:
        window = text[: limit + 1]
        cut = -1

        for sep in ("\n\n", "\n"):
            idx = window.rfind(sep)
            if idx >= min_cut:
                cut = idx
                break

        if cut == -1:
            ends = [m.end() for m in re.finditer(r"[.!?…。](\s|$)", window)]
            ends = [e for e in ends if e >= min_cut]
            if ends:
                cut = ends[-1]

        if cut == -1:
            idx = window.rfind(" ")
            cut = idx if idx > 0 else limit

        head = text[:cut].rstrip()
        last_nl = head.rfind("\n")
        if last_nl >= min_cut and _is_header(head[last_nl + 1 :]):
            cut = last_nl

        chunk = text[:cut].strip()
        if not chunk:
            chunk = text[:limit]
            cut = limit
        parts.append(chunk)
        text = text[cut:].strip()

    if text:
        parts.append(text)
    return parts


SECTION_HEADER_RE = re.compile(r"^\s*(?![\-\*•▪◦✅✔➤►→])[^\n]{2,60}:\s*$")


def _is_header(line: str) -> bool:
    return bool(SECTION_HEADER_RE.match(line))


def _split_sections(text: str) -> list:
    """
    Tách text thành các KHỐI, trả về [(sep_trước, nội_dung)].
    Một khối mới bắt đầu khi gặp:
      - dòng trống (sep = "\\n\\n")
      - dòng tiêu đề mục kể cả khi KHÔNG có dòng trống phía trên (sep = "\\n")
    Nhờ vậy mỗi mục (tiêu đề + các dòng bên dưới) luôn đi cùng nhau, không bị cắt đôi
    giữa chừng, và khi ghép lại vẫn giữ đúng kiểu xuống dòng ban đầu.
    """
    blocks = []
    cur = []
    sep = ""
    for line in text.split("\n"):
        if not line.strip():
            if cur:
                blocks.append((sep, "\n".join(cur).strip()))
                cur = []
            sep = "\n\n" if blocks else ""
            continue
        if cur and _is_header(line) and not _is_header(cur[-1]):
            blocks.append((sep, "\n".join(cur).strip()))
            cur = []
            sep = "\n"
        cur.append(line)
    if cur:
        blocks.append((sep, "\n".join(cur).strip()))
    return [(sp, bl) for sp, bl in blocks if bl]


def split_by_blocks(
    text: str, first_limit: int, limit: int, rebalance_tail: bool = True
) -> list:
    """
    Chia text thành nhiều phần theo KHỐI/MỤC (xem _split_sections).
    - Phần đầu tối đa first_limit ký tự, các phần sau tối đa limit ký tự.
    - Xếp khối nguyên vẹn vào phần hiện tại cho tới khi đầy; khối nào không vừa
      thì sang phần kế tiếp.
    - Chỉ khi 1 khối tự nó dài hơn giới hạn mới cắt bên trong khối (split_text).
    - rebalance_tail=True: phần cuối quá ngắn (< MIN_TAIL_CHARS) được kéo thêm khối
      cuối của phần trước. Đặt False cho BÀI CHÍNH để bài chính luôn được lấp đầy
      tối đa (không bị rút bớt xuống comment).
    """
    text = (text or "").strip()
    if not text:
        return []

    parts = []
    cur = ""

    for sep, block in _split_sections(text):
        lim = first_limit if not parts else limit

        if cur and len(cur) + len(sep) + len(block) <= lim:
            cur += sep + block
            continue
        if not cur and len(block) <= lim:
            cur = block
            continue

        if cur:
            parts.append(cur)
            cur = ""
            lim = limit
            if len(block) <= lim:
                cur = block
                continue

        pieces = split_text(block, lim)
        parts.extend(pieces[:-1])
        cur = pieces[-1]

    if cur:
        parts.append(cur)

    if rebalance_tail:
        while len(parts) >= 2 and len(parts[-1]) < MIN_TAIL_CHARS:
            prev_blocks = parts[-2].split("\n\n")
            if len(prev_blocks) < 2:
                break
            new_tail = prev_blocks[-1] + "\n\n" + parts[-1]
            if len(new_tail) > limit:
                break
            parts[-2] = "\n\n".join(prev_blocks[:-1])
            parts[-1] = new_tail

    return parts


def split_main_and_overflow(job_content: str, topic: str):
    """
    Trả về (bài chính, [các đoạn dư]).
    Bài chính được lấp đầy tối đa (chừa chỗ cho hashtag topic gõ thêm sau "\\n\\n#");
    phần còn lại là phần dư sẽ xuống comment.
    """
    first_limit = THREADS_CHAR_LIMIT
    clean_topic = topic.replace("#", "").strip() if topic else ""
    if clean_topic:
        first_limit -= len(clean_topic) + 3

    parts = split_by_blocks(
        job_content, first_limit, THREADS_CHAR_LIMIT, rebalance_tail=False
    )
    if not parts:
        return "", []
    return parts[0], parts[1:]


def build_reply_parts(overflow_parts: list, thread_content: str) -> list:
    """
    Tạo danh sách comment:
    - MERGE bật & có phần dư: nối phần dư + Thread Content thành 1 mạch rồi xếp theo mục;
      comment đầu còn chỗ thì Thread Content được đưa vào cùng comment đó.
    - Ngược lại: phần dư là các comment riêng, Thread Content là comment riêng ở cuối.
    """
    thread_parts = split_by_blocks(thread_content, THREADS_CHAR_LIMIT, THREADS_CHAR_LIMIT)
    if MERGE_OVERFLOW_WITH_THREAD and overflow_parts:
        combined = "\n\n".join(overflow_parts + ([thread_content] if thread_content else []))
        return split_by_blocks(combined, THREADS_CHAR_LIMIT, THREADS_CHAR_LIMIT)
    return list(overflow_parts) + thread_parts


def convert_google_drive(url: str) -> str:
    match = re.search(r"/d/([^/]+)/", url)
    if match:
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
    match = re.search(r"id=([^&]+)", url)
    if match:
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
    return url


def get_filename_from_response(response, default="image.jpg"):
    cd = response.headers.get("Content-Disposition", "")
    if "filename=" in cd:
        return cd.split("filename=")[-1].strip('"')
    return default


def make_square(image_path, min_size=1080, fill_color=(0, 0, 0)):
    img = Image.open(image_path)
    w, h = img.size
    size = max(min_size, w, h)
    new_img = Image.new("RGB", (size, size), fill_color)
    new_img.paste(img, ((size - w) // 2, (size - h) // 2))
    new_img.save(image_path)


def download_image(url, folder="tmp_images"):
    folder_path = os.path.join(BASE_DIR, folder)
    os.makedirs(folder_path, exist_ok=True)
    url = convert_google_drive(url)
    response = requests.get(url, allow_redirects=True, timeout=20)
    if response.status_code != 200:
        raise Exception(f"Image download failed: {response.status_code}")
    filename = get_filename_from_response(response).replace("/", "_")
    full_path = os.path.join(folder_path, filename)
    with open(full_path, "wb") as f:
        f.write(response.content)
    make_square(full_path)
    return full_path


# =========================
# GOOGLE SHEET LOGIC
# =========================
_gc = None
_sa_email = ""
_ws_cache = {}  # tab_name -> worksheet
_header_cache = {}  # tab_name -> {col_name: index}


def _reset_sheet_cache():
    global _gc
    _gc = None
    _ws_cache.clear()
    _header_cache.clear()


def _api_status(e: Exception):
    try:
        return e.response.status_code
    except Exception:
        return None


def _is_transient_error(e: Exception) -> bool:
    if isinstance(e, (requests.exceptions.RequestException, TransportError)):
        return True
    if isinstance(e, (ConnectionError, TimeoutError)):
        return True
    if isinstance(e, gspread.exceptions.APIError):
        return _api_status(e) in (429, 500, 502, 503, 504)
    return False


def retry_network(func):
    """Tự thử lại khi lỗi mạng tạm thời (vd: WinError 10054) với thời gian chờ tăng dần."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        for attempt in range(1, SHEET_MAX_RETRIES + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                if not _is_transient_error(e) or attempt == SHEET_MAX_RETRIES:
                    raise
                wait = SHEET_RETRY_BASE_DELAY * (2 ** (attempt - 1))
                print(
                    f"⚠ Lỗi mạng khi gọi Google ({func.__name__}, lần {attempt}/{SHEET_MAX_RETRIES}): "
                    f"{type(e).__name__}. Thử lại sau {wait}s..."
                )
                _reset_sheet_cache()
                time.sleep(wait)

    return wrapper


def get_client():
    global _gc, _sa_email
    if _gc is None:
        creds = Credentials.from_service_account_file(
            CREDENTIAL_FILE, scopes=SHEET_SCOPES
        )
        _sa_email = getattr(creds, "service_account_email", "") or ""
        _gc = gspread.authorize(creds)
        _gc.set_timeout(SHEET_HTTP_TIMEOUT)
    return _gc


@retry_network
def connect_sheet(tab_name):
    if tab_name not in _ws_cache:
        _ws_cache[tab_name] = get_client().open_by_url(RECRUIT_SHEET_URL).worksheet(
            tab_name
        )
    return _ws_cache[tab_name]


def _col_index(tab_name: str, col_name: str) -> int:
    if tab_name not in _header_cache:
        sheet = connect_sheet(tab_name)
        headers = sheet.row_values(1)
        _header_cache[tab_name] = {h.strip(): i for i, h in enumerate(headers, start=1)}
    idx = _header_cache[tab_name].get(col_name)
    if idx is None:
        raise Exception(f"❌ Không tìm thấy cột: {col_name} trong tab {tab_name}")
    return idx


@retry_network
def _probe_sheet_write():
    """Ghi lại đúng giá trị cũ vào ô Posted của dòng 2 → kiểm tra quyền GHI mà không đổi dữ liệu."""
    sheet = connect_sheet(RECRUIT_TAB_NAME)
    col = _col_index(RECRUIT_TAB_NAME, COL_POSTED)
    current = sheet.cell(2, col).value or ""
    sheet.update_cell(2, col, current)


def check_sheet_write_access() -> bool:
    """Chạy trước khi đăng bài: nếu không ghi được sheet thì dừng, tránh đăng xong mà không lưu trạng thái."""
    try:
        _probe_sheet_write()
        return True
    except gspread.exceptions.APIError as e:
        if _api_status(e) == 403:
            print("⛔ Service account chỉ có quyền ĐỌC, chưa có quyền GHI vào Google Sheet.")
            print(
                f"   → Mở Google Sheet → Share → thêm email: {_sa_email or '(xem client_email trong credentials.json)'} "
                "với quyền 'Editor' (Người chỉnh sửa)."
            )
            print(
                "   → Nếu đã là Editor: kiểm tra Data → 'Protected sheets and ranges' "
                f"xem cột {COL_POSTED}/{COL_LINK_POST}/{COL_DATE} có bị khóa không."
            )
        else:
            print(f"⛔ Không ghi thử được lên sheet: {type(e).__name__}: {e}")
        return False
    except Exception as e:
        print(f"⛔ Không ghi thử được lên sheet: {type(e).__name__}: {e}")
        return False


@retry_network
def get_all_accounts():
    sheet = connect_sheet(ACCOUNT_TAB_NAME)
    all_values = sheet.get_all_values()
    if len(all_values) < 2:
        return {}

    headers = all_values[0]
    records = [dict(zip(headers, row)) for row in all_values[1:]]

    accounts = {}
    for row in records:
        code = str(row.get("AccountsCode", "")).strip()
        if code:
            active_raw = str(row.get(COL_ACC_ACTIVE, "")).strip().upper()
            accounts[code] = {
                "email": str(row.get("Email", "")).strip(),
                "password": str(row.get("Password", "")).strip(),
                "username": str(row.get(COL_ACC_USERNAME, "")).strip().lstrip("@"),
                "active": active_raw not in INACTIVE_VALUES,
            }
    return accounts


@retry_network
def get_unposted_rows(limit=MAX_POSTS_PER_RUN):
    sheet = connect_sheet(RECRUIT_TAB_NAME)
    all_values = sheet.get_all_values()
    if len(all_values) < 2:
        return []

    headers = all_values[0]
    rows = [dict(zip(headers, row)) for row in all_values[1:]]

    results = []
    for idx, row in enumerate(rows, start=2):
        posted = str(row.get(COL_POSTED, "")).strip().upper()
        if posted == STATUS_YES:
            if str(row.get(COL_LINK_POST, "")).strip():
                continue
            posted = STATUS_YES_NO_LINK
        results.append({"row_index": idx, "data": row, "status": posted})
        if len(results) >= limit:
            break
    return results


def _now_vn() -> str:
    tz_vn = pytz.timezone("Asia/Ho_Chi_Minh")
    return datetime.now(tz_vn).strftime("%Y-%m-%d %H:%M:%S")


@retry_network
def _write_row_status(row_index: int, status: str, link_or_msg: str):
    """Ghi 3 ô (Posted / Link post / Date) trong 1 request duy nhất."""
    sheet = connect_sheet(RECRUIT_TAB_NAME)
    updates = [
        {
            "range": rowcol_to_a1(row_index, _col_index(RECRUIT_TAB_NAME, COL_POSTED)),
            "values": [[status]],
        },
        {
            "range": rowcol_to_a1(
                row_index, _col_index(RECRUIT_TAB_NAME, COL_LINK_POST)
            ),
            "values": [[link_or_msg]],
        },
        {
            "range": rowcol_to_a1(row_index, _col_index(RECRUIT_TAB_NAME, COL_DATE)),
            "values": [[_now_vn()]],
        },
    ]
    sheet.batch_update(updates)


def _log_unsynced(row_index: int, status: str, link_or_msg: str):
    try:
        with open(UNSYNCED_FILE, "a", encoding="utf-8") as f:
            f.write(
                json.dumps(
                    {
                        "time": _now_vn(),
                        "row_index": row_index,
                        "status": status,
                        "value": link_or_msg,
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
        print(f"📝 Đã lưu tạm trạng thái vào {UNSYNCED_FILE} để bạn cập nhật sheet thủ công.")
    except Exception as e:
        print(f"⚠ Không ghi được file tạm {UNSYNCED_FILE}: {e}")


def _safe_write_status(row_index: int, status: str, link_or_msg: str) -> bool:
    """Không bao giờ raise: lỗi ghi sheet chỉ được log + lưu file tạm."""
    try:
        _write_row_status(row_index, status, link_or_msg)
        return True
    except Exception as e:
        print(
            f"⚠ KHÔNG ghi được trạng thái {status} cho dòng {row_index} lên sheet: "
            f"{type(e).__name__}: {e}"
        )
        _log_unsynced(row_index, status, link_or_msg)
        return False


def mark_posted(row_index: int, post_url: str):
    return _safe_write_status(row_index, STATUS_YES, post_url)


def mark_error(row_index: int, error_msg: str):
    return _safe_write_status(row_index, STATUS_ERROR, error_msg)


def mark_reply_error(row_index: int, post_url: str):
    """Bài chính đã lên, comment phụ lỗi: giữ link bài để lần sau chỉ thử lại comment."""
    return _safe_write_status(row_index, STATUS_REPLY_ERROR, post_url)


def mark_check(row_index: int, msg: str):
    """Đã bấm đăng nhưng không xác minh được: đánh dấu để kiểm tra tay, KHÔNG tự đăng lại."""
    return _safe_write_status(row_index, STATUS_CHECK, msg)


def _load_disabled_cache() -> dict:
    """{MÃ_VIẾT_HOA: epoch_giây_lúc_phát_hiện}; lỗi đọc file thì coi như chưa có."""
    try:
        with open(DISABLED_CACHE_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return {str(k).upper(): float(v) for k, v in data.items()}
    except Exception:
        return {}


def _save_disabled_cache(cache: dict):
    try:
        with open(DISABLED_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠ Không ghi được {DISABLED_CACHE_FILE}: {e}")


def recently_disabled_codes() -> set:
    """Các mã tài khoản được phát hiện bị khóa trong vòng DISABLED_RECHECK_HOURS giờ gần đây."""
    ttl = DISABLED_RECHECK_HOURS * 3600
    now = time.time()
    return {code for code, ts in _load_disabled_cache().items() if now - ts < ttl}


def remember_disabled(account_code: str):
    cache = _load_disabled_cache()
    cache[account_code.strip().upper()] = time.time()
    _save_disabled_cache(cache)


def mark_account_disabled_row(row_index: int, account_code: str):
    """Tài khoản bị khóa: ghi 'Account disabled' vào cột Posted (không phải YES nên dòng vẫn chờ đăng)."""
    return _safe_write_status(
        row_index,
        STATUS_ACCOUNT_DISABLED,
        f"Tài khoản {account_code} bị Threads vô hiệu hóa. Hãy đổi AccountsCode.",
    )


def get_existing_post_url(data: dict) -> str:
    """Lấy link bài chính đã lưu ở dòng REPLY_ERROR (nếu hợp lệ)."""
    url = str(data.get(COL_LINK_POST, "")).strip()
    return url if url.startswith("http") and "/post/" in url else ""


# ==========================================
# JS HELPERS
# ==========================================
PROFILE_PATH_JS = """
() => {
  const labelRe = /trang cá nhân|profile|hồ sơ/i;
  const isProfilePath = (p) => /^\\/@[^\\/]+\\/?$/.test(p);
  const anchors = [...document.querySelectorAll('a[href*="/@"]')]
    .filter(a => isProfilePath(a.pathname));

  const byLabel = anchors.find(a => {
    if (labelRe.test(a.getAttribute('aria-label') || '')) return true;
    return [...a.querySelectorAll('svg[aria-label]')]
      .some(s => labelRe.test(s.getAttribute('aria-label') || ''));
  });
  if (byLabel) return byLabel.pathname;

  const rail = anchors.filter(a => {
    const r = a.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && r.left < 150 && (a.innerText || '').trim() === '';
  });
  if (rail.length) return rail[0].pathname;
  return '';
}
"""

PROFILE_POSTS_JS = """
(prefix) => {
  const want = prefix.replace(/^\\//, '').toLowerCase();
  const out = [];
  for (const a of document.querySelectorAll('a[href*="/post/"]')) {
    const parts = a.pathname.split('/').filter(Boolean);
    if (parts.length === 3 && parts[1] === 'post' && parts[0].toLowerCase() === want) {
      if (!out.includes(a.pathname)) out.push(a.pathname);
    }
  }
  return out;
}
"""

DEBUG_INFO_JS = """
() => ({
  url: location.href,
  title: document.title,
  svgLabels: [...new Set([...document.querySelectorAll('svg[aria-label]')]
    .map(s => s.getAttribute('aria-label')))].slice(0, 30),
  ariaLabels: [...new Set([...document.querySelectorAll('[aria-label]')]
    .map(e => e.tagName.toLowerCase() + ':' + e.getAttribute('aria-label')))].slice(0, 40),
  buttons: [...new Set([...document.querySelectorAll('[role="button"]')]
    .map(e => (e.innerText || '').trim().slice(0, 30))
    .filter(Boolean))].slice(0, 30),
  hasEditable: document.querySelectorAll('div[contenteditable="true"]').length,
  links: [...new Set([...document.querySelectorAll('a[href]')]
    .map(a => a.getAttribute('href'))
    .filter(h => h && h.length < 80))].slice(0, 40),
  text: (document.body.innerText || '').slice(0, 300)
})
"""


# ==========================================
# THREADS BOT (PLAYWRIGHT)
# ==========================================
class ThreadsBot:
    def __init__(
        self,
        account_code: str,
        email: str,
        password: str,
        headless=True,
        username: str = "",
    ):
        self.headless = headless
        self.account_code = account_code
        self.email = email
        self.password = password
        self.username = (username or "").strip().lstrip("@")
        self.profile_path = ""

        self.cookie_file = os.path.join(
            BASE_DIR, "cookies", f"{self.account_code}.json"
        )
        os.makedirs(os.path.dirname(self.cookie_file), exist_ok=True)

        self.error_dir = os.path.join(BASE_DIR, "error_logs")
        os.makedirs(self.error_dir, exist_ok=True)

        self.pw = None
        self.browser = None
        self.context = None
        self.page = None

    @property
    def profile_url(self) -> str:
        return f"{THREADS_URL}{self.profile_path}"

    async def _pause(self, rng):
        await self.page.wait_for_timeout(int(random.uniform(*rng) * 1000))

    async def _settle(self, extra_ms: int = 2500):
        """Chờ trang tải xong 'vừa đủ' (không dùng networkidle vì Threads luôn có request nền)."""
        try:
            await self.page.wait_for_load_state("load", timeout=10000)
        except Exception:
            pass
        await self.page.wait_for_timeout(extra_ms)

    async def _goto(self, url: str, attempts: int = GOTO_ATTEMPTS):
        last_error = None
        for n in range(1, attempts + 1):
            try:
                await self.page.goto(
                    url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS
                )
                await self._settle()
                return
            except Exception as e:
                last_error = e
                print(f"   ⚠ Mở {url} lỗi (lần {n}/{attempts}): {type(e).__name__}")
                await self.page.wait_for_timeout(3000)
        raise last_error

    async def _dump_debug(self, tag: str):
        """Lưu ảnh + HTML + thông tin trang để chẩn đoán khi không tìm thấy phần tử."""
        base = os.path.join(self.error_dir, f"error_{tag}_{self.account_code}")
        try:
            await self.page.screenshot(path=base + ".png")
        except Exception:
            pass
        try:
            html = await self.page.content()
            with open(base + ".html", "w", encoding="utf-8") as f:
                f.write(html)
        except Exception:
            pass
        try:
            info = await self.page.evaluate(DEBUG_INFO_JS)
            print(f"🧪 DEBUG [{tag}] {self.account_code}:")
            print(json.dumps(info, ensure_ascii=False, indent=2))
        except Exception as e:
            print(f"🧪 DEBUG [{tag}] không đọc được thông tin trang: {e}")
        print(f"   Đã lưu {base}.png / .html")

    async def start(self):
        self.pw = await async_playwright().start()
        self.browser = await self.pw.chromium.launch(
            headless=self.headless,
            channel="chrome",
            args=["--disable-blink-features=AutomationControlled", "--lang=vi-VN"],
        )

        context_options = {
            "viewport": {"width": 1280, "height": 900},
            "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
            "locale": "vi-VN",
            "timezone_id": "Asia/Ho_Chi_Minh",
        }

        if os.path.exists(self.cookie_file):
            context_options["storage_state"] = self.cookie_file
            print(f"🍪 Đã tìm thấy Cookie cho {self.account_code}. Đang nạp...")

        self.context = await self.browser.new_context(**context_options)
        self.context.set_default_timeout(DEFAULT_TIMEOUT_MS)
        self.context.set_default_navigation_timeout(NAV_TIMEOUT_MS)
        self.page = await self.context.new_page()

        await self._goto(THREADS_URL)
        await self._raise_if_disabled()
        is_logged_in = await self._is_logged_in()

        if not is_logged_in:
            print(
                f"⚠ Cookie hỏng hoặc chưa có cho {self.account_code}. Tiến hành auto-login..."
            )
            try:
                await self._login()
            except Exception:
                await self._raise_if_disabled()
                raise
            await self._raise_if_disabled()
        else:
            print(f"✅ Tài khoản {self.account_code} đăng nhập thành công bằng Cookie.")

    async def _raise_if_disabled(self):
        """Phát hiện trang 'Chúng tôi đã vô hiệu hóa tài khoản của bạn'."""
        try:
            await self.page.wait_for_timeout(2000)
            body_text = (await self.page.inner_text("body")).lower()
        except Exception:
            return

        if any(marker in body_text for marker in DISABLED_MARKERS):
            await self.page.screenshot(
                path=os.path.join(
                    self.error_dir, f"account_disabled_{self.account_code}.png"
                )
            )
            raise AccountDisabledError(
                f"Tài khoản {self.account_code} bị Threads vô hiệu hóa/khóa"
            )

    async def close(self):
        try:
            if self.context:
                await self.context.close()
        except Exception:
            pass
        try:
            if self.browser:
                await self.browser.close()
        except Exception:
            pass
        try:
            if self.pw:
                await self.pw.stop()
        except Exception:
            pass

    async def _is_logged_in(self) -> bool:
        try:
            await self.page.wait_for_timeout(3000)
            if "/login" in self.page.url:
                return False

            guest_texts = [
                "Đăng nhập hoặc đăng ký",
                "Đăng nhập bằng tên người dùng",
                "Log in with phone",
            ]
            for t in guest_texts:
                if await self.page.get_by_text(t, exact=True).first.is_visible():
                    return False
            return True
        except Exception:
            return False

    async def _click_first_visible_text(self, texts, timeout_ms: int = 5000) -> bool:
        end = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < end:
            for t in texts:
                try:
                    loc = self.page.get_by_text(t, exact=True).first
                    if await loc.is_visible():
                        await loc.click(force=True)
                        return True
                except Exception:
                    pass
            await self.page.wait_for_timeout(500)
        return False

    async def _login(self):
        await self._goto(f"{THREADS_URL}/login")
        try:
            await self.page.wait_for_timeout(3000)
            try:
                clicked = await self._click_first_visible_text(
                    [
                        "Đăng nhập bằng tên người dùng",
                        "Log in with phone",
                        "Log in with email",
                    ],
                    timeout_ms=5000,
                )
                if clicked:
                    print("🔄 Đang click mở Form gõ Email/Password...")
                    await self.page.wait_for_timeout(2000)
            except Exception:
                pass

            await self.page.wait_for_selector(
                'input[type="text"], input[name="username"]', timeout=10000
            )
            await self.page.fill(
                'input[type="text"], input[name="username"]', self.email
            )
            await self.page.wait_for_timeout(500)

            await self.page.fill(
                'input[type="password"], input[name="password"]', self.password
            )
            await self.page.wait_for_timeout(500)

            print("🔄 Đang nhấn phím Enter để Đăng nhập...")
            await self.page.keyboard.press("Enter")
            await self.page.wait_for_timeout(6000)

            await self.page.wait_for_timeout(5000)
            still_guest = False
            if await self.page.get_by_text(
                "Đăng nhập bằng tên người dùng", exact=True
            ).first.is_visible():
                still_guest = True
            elif await self.page.locator(
                "button:has-text('Đăng nhập'), button:has-text('Log in')"
            ).first.is_visible():
                still_guest = True
            if still_guest:
                raise Exception(
                    "Kẹt ở màn hình đăng nhập (Bị vướng Captcha hoặc Checkpoint). Hãy dùng script get_cookie trên máy tính."
                )

            print(f"✅ Đăng nhập xong cho {self.account_code}!")
            await self.context.storage_state(path=self.cookie_file)
            print(f"💾 Đã lưu Cookie mới vào: {self.cookie_file}")

            await self._goto(THREADS_URL)
            await self._dismiss_popups()

        except Exception as e:
            await self.page.screenshot(
                path=os.path.join(
                    self.error_dir, f"error_login_{self.account_code}.png"
                )
            )
            raise Exception(f"❌ Login tự động thất bại. Lỗi: {e}")

    async def _dismiss_popups(self):
        """Đóng các popup thường gặp (best-effort, lỗi thì bỏ qua)."""
        try:
            await self._click_first_visible_text(
                ["Lúc khác", "Not now", "Không phải bây giờ", "Để sau"],
                timeout_ms=2500,
            )
        except Exception:
            pass

    async def _resolve_profile_path(self, timeout_s: int = 20) -> str:
        """
        Xác định /@username của tài khoản.
        1) Cột Username ở tab Accounts (nếu có)
        2) Đọc URL sau khi mở nút 'Trang cá nhân/Profile'
        3) Dò link profile trên thanh menu trái (nhiều cách dự phòng)
        """
        if self.profile_path:
            return self.profile_path

        if self.username:
            self.profile_path = f"/@{self.username}"
            print(f"👤 Dùng Username từ sheet: {self.profile_path}")
            return self.profile_path

        await self._goto(THREADS_URL)
        await self._dismiss_popups()

        def profile_path_from_url() -> str:
            match = re.search(r"https?://[^/]+(/@[^/?#]+)", self.page.url)
            return match.group(1).rstrip("/") if match else ""

        async def open_profile_button() -> str:
            try:
                button = self.page.get_by_role(
                    "button",
                    name=re.compile(r"(Trang cá nhân|Profile)", re.IGNORECASE),
                ).first
                if await button.is_visible():
                    await button.click(force=True)
                    await self.page.wait_for_timeout(1500)
                    return profile_path_from_url()
            except Exception:
                pass
            return ""

        path = profile_path_from_url()
        if not path:
            path = await open_profile_button()
        if path:
            self.profile_path = path
            print(f"👤 Trang cá nhân: {self.profile_path}")
            return self.profile_path

        end = time.monotonic() + timeout_s
        while time.monotonic() < end:
            try:
                path = profile_path_from_url() or await self.page.evaluate(
                    PROFILE_PATH_JS
                )
            except Exception:
                path = ""
            if path:
                self.profile_path = path.rstrip("/")
                print(f"👤 Trang cá nhân: {self.profile_path}")
                return self.profile_path
            await self.page.wait_for_timeout(1000)

        await self._dump_debug("get_profile")
        return ""

    async def _profile_post_urls(self, wait_s: int = 12) -> list:
        """Mở trang cá nhân và trả về danh sách link bài (mới nhất trước)."""
        await self._goto(self.profile_url)
        urls = []
        end = time.monotonic() + wait_s
        while time.monotonic() < end:
            try:
                paths = await self.page.evaluate(PROFILE_POSTS_JS, self.profile_path)
            except Exception:
                paths = []
            if paths:
                urls = [f"{THREADS_URL}{p}" for p in paths]
                break
            await self.page.wait_for_timeout(1000)
        return urls

    async def _wait_for_new_post(self, old_urls: set) -> str:
        for attempt in range(6):
            print(f"   ⏳ Tải lại trang cá nhân (Lần {attempt + 1}/6)...")
            urls = await self._profile_post_urls(wait_s=8)
            new_urls = [u for u in urls if u not in old_urls]
            if new_urls:
                return new_urls[0]
            await self.page.wait_for_timeout(3000)

        await self._dump_debug("missing_post")
        return ""

    async def post(self, text: str, image_path: str | None = None, topic: str = ""):
        if not text.strip():
            raise ValueError("❌ Nội dung bài post trống")

        if not await self._resolve_profile_path():
            raise Exception(
                "Không xác định được trang cá nhân → CHƯA đăng bài (xem error_logs/error_get_profile_*.png/.html). "
                f"Có thể thêm cột '{COL_ACC_USERNAME}' vào tab {ACCOUNT_TAB_NAME} để bỏ qua bước dò."
            )

        print("🔍 Đang ghi nhận danh sách bài cũ trên tường để đối chiếu...")
        old_post_urls = set(await self._profile_post_urls())

        await self._open_composer()
        await self._type_text(
            text, editor=self.page.locator("div[contenteditable='true']").first
        )

        if topic:
            clean_topic = topic.replace("#", "").strip()
            if clean_topic:
                print(f"🏷️ Đang gắn Chủ đề (Topic): {clean_topic}...")
                await self.page.keyboard.type(f"\n\n#{clean_topic}", delay=150)
                await self.page.wait_for_timeout(2000)
                await self.page.keyboard.press("Enter")
                await self.page.wait_for_timeout(1000)

        await self.page.wait_for_timeout(2000)

        if image_path:
            await self._upload_image(image_path)

        print("🚀 Đang bấm nút đăng bài...")
        await self._submit_post()

        print("🔍 Đang chờ Threads load xong bài viết mới...")
        post_url = await self._wait_for_new_post(old_post_urls)
        if not post_url:
            raise PostUnverifiedError(
                "Đã bấm Đăng nhưng không thấy bài mới trên profile. "
                "Hãy kiểm tra tài khoản thủ công (xem error_logs/error_missing_post_*)"
            )
        return post_url

    # ---------- comment ----------
    @staticmethod
    def _snippet_of(text: str) -> str:
        return text.strip().split("\n")[0][:25]

    async def _comment_exists(
        self, post_url: str, text: str, main_text: str = ""
    ) -> bool:
        """
        Kiểm tra comment (theo đoạn đầu nội dung) đã có dưới bài chưa.
        Trang bài cũng chứa BÀI CHÍNH, nên nếu đoạn đầu comment trùng với bài chính
        thì phải có ít nhất 2 chỗ khớp mới tính là comment đã tồn tại.
        """
        await self._goto(post_url)
        await self.page.wait_for_timeout(2000)
        snippet = self._snippet_of(text)
        if not snippet:
            return False

        baseline = 1 if snippet.lower() in (main_text or "").lower() else 0
        for _ in range(3):
            try:
                count = await self.page.get_by_text(snippet, exact=False).count()
            except Exception:
                count = 0
            if count > baseline:
                return True
            await self.page.mouse.wheel(0, 600)
            await self.page.wait_for_timeout(2000)
        return False

    async def reply_to_post(self, post_url: str, texts, main_text: str = ""):
        """
        Đăng lần lượt các comment phụ (phần dư của bài chính + Thread Content).
        - Mỗi comment thử lại tối đa REPLY_MAX_ATTEMPTS lần.
        - Comment nào đã có sẵn dưới bài thì bỏ qua (tránh đăng trùng khi chạy lại).
        - Trước mỗi lần thử lại, kiểm tra xem lần trước thực ra đã lên chưa.
        - Nếu vẫn lỗi thì RAISE để sheet ghi REPLY_ERROR.
        """
        if isinstance(texts, str):
            texts = [texts]
        texts = [t for t in texts if t and t.strip()]
        total = len(texts)

        for n, text in enumerate(texts, start=1):
            if await self._comment_exists(post_url, text, main_text):
                print(f"⏭ Comment {n}/{total} đã có dưới bài → bỏ qua.")
                continue

            last_error = None
            for attempt in range(1, REPLY_MAX_ATTEMPTS + 1):
                if attempt > 1 and await self._comment_exists(
                    post_url, text, main_text
                ):
                    print(f"✅ Comment {n}/{total} thực ra đã lên ở lần trước.")
                    last_error = None
                    break

                print(
                    f"💬 Đang đăng comment {n}/{total} - lần {attempt}/{REPLY_MAX_ATTEMPTS}..."
                )
                try:
                    await self._reply_once(post_url, text, main_text)
                    print(f"✅ Đã đăng comment {n}/{total} thành công!")
                    last_error = None
                    break
                except AccountDisabledError:
                    raise
                except Exception as e:
                    last_error = e
                    try:
                        await self.page.screenshot(
                            path=os.path.join(
                                self.error_dir,
                                f"error_reply_{self.account_code}_c{n}_try{attempt}.png",
                            )
                        )
                    except Exception:
                        pass
                    print(f"⚠ Comment {n} lần {attempt} lỗi: {e}")
                    try:
                        await self.page.keyboard.press("Escape")
                    except Exception:
                        pass
                    await self.page.wait_for_timeout(random.randint(3000, 5000))

            if last_error is not None:
                raise Exception(
                    f"Không thể đăng comment {n}/{total} sau {REPLY_MAX_ATTEMPTS} lần: {last_error}"
                )

            if n < total:
                await self.page.wait_for_timeout(random.randint(4000, 7000))

    async def _find_reply_button(self, timeout_ms: int = 20000):
        """Tìm nút Trả lời của bài chính bằng nhiều cách; trả về locator hiển thị được hoặc None."""
        strategies = [
            lambda: self.page.locator(
                "div[role='button']:has(svg[aria-label='Trả lời']), "
                "div[role='button']:has(svg[aria-label='Reply'])"
            ),
            lambda: self.page.locator("[aria-label='Trả lời'], [aria-label='Reply']"),
            lambda: self.page.locator(
                "svg[aria-label*='Trả lời'], svg[aria-label*='Reply']"
            ),
            lambda: self.page.get_by_role("button", name="Trả lời", exact=True),
            lambda: self.page.get_by_role("button", name="Reply", exact=True),
        ]
        end = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < end:
            for make in strategies:
                try:
                    loc = make()
                    for i in range(min(await loc.count(), 5)):
                        item = loc.nth(i)
                        if await item.is_visible():
                            return item
                except Exception:
                    pass
            await self.page.wait_for_timeout(1000)
        return None

    async def _open_reply_editor(self):
        """
        Mở khung soạn comment. Thứ tự:
        1) Bấm nút Trả lời
        2) Dự phòng: ô soạn sẵn trên trang ("Thêm vào thread"/"Add to thread")
        3) Dự phòng: contenteditable đang hiển thị trên trang
        """
        reply_btn = await self._find_reply_button()
        if reply_btn is not None:
            await reply_btn.scroll_into_view_if_needed()
            await reply_btn.click(force=True)
            await self.page.wait_for_timeout(2000)
            return

        print("   ⚠ Không thấy nút Trả lời → thử ô soạn sẵn trên trang...")
        for t in ["Thêm vào thread", "Add to thread"]:
            try:
                loc = self.page.get_by_text(t, exact=False).first
                if await loc.is_visible():
                    await loc.click(force=True)
                    await self.page.wait_for_timeout(2000)
                    return
            except Exception:
                pass

        try:
            editable = self.page.locator("div[contenteditable='true']").last
            if await editable.is_visible():
                await editable.click(force=True)
                await self.page.wait_for_timeout(1500)
                return
        except Exception:
            pass

        await self._dump_debug("reply_button")
        raise Exception("Không tìm thấy nút Trả lời / ô soạn comment trên trang bài viết")

    async def _reply_once(self, post_url: str, text: str, main_text: str = ""):
        await self._goto(post_url)
        await self._dismiss_popups()
        await self._raise_if_disabled()
        await self.page.wait_for_timeout(random.randint(2000, 4000))

        await self._open_reply_editor()

        dialog = self.page.locator("div[role='dialog']").last
        if await dialog.count() > 0:
            editor = dialog.locator("div[contenteditable='true']").first
        else:
            editor = self.page.locator("div[contenteditable='true']").last
        await editor.wait_for(state="visible", timeout=10000)
        await editor.click(force=True)

        await self._type_text(text, editor=editor)
        await self.page.wait_for_timeout(1000)

        await self._submit_post(scope="dialog")
        await self.page.wait_for_timeout(4000)

        if not await self._comment_exists(post_url, text, main_text):
            raise Exception(
                "Gửi comment xong nhưng không thấy comment xuất hiện dưới bài"
            )

    async def _open_composer(self):
        try:
            await self.page.keyboard.press("Escape")
            await self.page.wait_for_timeout(500)
            await self.page.keyboard.press("Escape")
            await self.page.wait_for_timeout(1000)

            try:
                nav_btn = self.page.locator("a[href='/compose']").first
                await nav_btn.wait_for(state="visible", timeout=3000)
                await nav_btn.click(force=True)
            except Exception:
                try:
                    trigger_vi = self.page.locator("text='Có gì mới?'").first
                    await trigger_vi.wait_for(state="visible", timeout=3000)
                    await trigger_vi.click(force=True)
                except Exception:
                    try:
                        trigger_en = self.page.locator('text="What\'s new?"').first
                        await trigger_en.wait_for(state="visible", timeout=3000)
                        await trigger_en.click(force=True)
                    except Exception:
                        plus_btn = self.page.locator(
                            "svg[aria-label='Tạo'], svg[aria-label='Create'], svg[aria-label='Bắt đầu thread mới']"
                        ).first
                        await plus_btn.click(force=True, timeout=5000)

            await self.page.wait_for_timeout(2000)
            editor = self.page.locator("div[contenteditable='true']").first
            await editor.wait_for(state="visible", timeout=15000)
            await editor.click(force=True)
            await self._pause(POST_DELAY_RANGE)

        except Exception as e:
            await self.page.screenshot(
                path=os.path.join(
                    self.error_dir, f"error_open_composer_{self.account_code}.png"
                )
            )
            raise Exception(f"Không thể mở khung đăng bài: {e}")

    @staticmethod
    def _compact(s: str) -> str:
        """Bỏ mọi khoảng trắng/ký tự vô hình để so sánh nội dung không phụ thuộc cách Threads xuống dòng."""
        s = unicodedata.normalize("NFC", s or "")
        return re.sub(r"[\s\u200b\u200c\u200d\u2060\ufeff\ufe0e\ufe0f]+", "", s)

    async def _read_focused_editor(self, editor=None):
        """Đọc nội dung khung soạn đang có con trỏ (nơi phím được gõ vào). None nếu mất focus."""
        try:
            return await self.page.evaluate(
                """() => {
                  const e = document.activeElement;
                  const ed = e && e.closest ? e.closest('[contenteditable="true"]') : null;
                  return ed ? ed.innerText : null;
                }"""
            )
        except Exception:
            return None

    async def _fill_editor(self, text: str):
        """
        Nhập từng DÒNG bằng insert_text (1 sự kiện/dòng) và nhấn Enter giữa các dòng.
        Không gõ từng phím: gõ từng ký tự khiến Threads dựng lại khung soạn khi nhận ra link /
        xem trước link, làm mất ký tự (lỗi comment chỉ còn '.com/oaVR ... Ứng tuy').
        """
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if line:
                await self.page.keyboard.insert_text(line)
            if i < len(lines) - 1:
                await self.page.keyboard.press("Enter")
            await self.page.wait_for_timeout(random.randint(150, 400))

    async def _type_text(self, text: str, editor=None):
        expected = self._compact(text)
        actual = ""
        for attempt in range(1, TYPE_MAX_ATTEMPTS + 1):
            if editor is not None:
                try:
                    await editor.click(force=True)
                except Exception:
                    pass

            await self._fill_editor(text)
            await self.page.wait_for_timeout(2000)

            if not VERIFY_TYPED_TEXT:
                await self._pause(POST_DELAY_RANGE)
                return

            raw = await self._read_focused_editor()
            actual = self._compact(raw) if raw is not None else ""
            if actual == expected:
                await self._pause(POST_DELAY_RANGE)
                return

            print(
                f"   ⚠ Nội dung trong khung soạn KHÔNG khớp (lần {attempt}/{TYPE_MAX_ATTEMPTS}): "
                f"mong đợi {len(expected)} ký tự, thực tế {len(actual)}. "
                f"Đầu: {actual[:25]!r} | Cuối: {actual[-25:]!r}"
            )
            try:
                await self.page.keyboard.press("Control+A")
                await self.page.keyboard.press("Backspace")
            except Exception:
                pass
            await self.page.wait_for_timeout(1500)

        await self._dump_debug("typed_mismatch")
        raise Exception(
            f"Gõ nội dung vào khung soạn bị lệch sau {TYPE_MAX_ATTEMPTS} lần "
            f"(mong đợi {len(expected)} ký tự, thực tế {len(actual)}) → KHÔNG đăng."
        )

    async def _upload_image(self, image_path: str):
        file_input = self.page.locator("input[type='file']").first
        await file_input.set_input_files(image_path)
        await self.page.wait_for_timeout(5000)

    async def _submit_post(self, scope: str = "page"):
        try:
            editor = self.page.locator("div[contenteditable='true']").last
            await editor.click(force=True, timeout=2000)
            await self.page.wait_for_timeout(500)
        except Exception:
            pass

        try:
            root = self.page
            if scope == "dialog":
                dialog = self.page.locator("div[role='dialog']").last
                if await dialog.count() > 0:
                    root = dialog
            post_btn = (
                root.locator("div[role='button']")
                .filter(has_text=re.compile(r"^(Đăng|Post)$"))
                .first
            )
            await post_btn.scroll_into_view_if_needed(timeout=2000)
            await post_btn.click(force=True, timeout=3000)
        except Exception:
            await self.page.keyboard.down("Control")
            await self.page.keyboard.press("Enter")
            await self.page.keyboard.up("Control")

        await self.page.wait_for_timeout(8000)


# ==========================================
# MAIN WORKFLOW
# ==========================================
async def run():
    print("🚀 START THREADS AUTO POST")
    try:
        accounts_dict = get_all_accounts()
        print(f"🔑 Đã load {len(accounts_dict)} tài khoản từ hệ thống.")
    except Exception as e:
        print(f"❌ Lỗi khi đọc tab Accounts: {e}")
        return

    if not check_sheet_write_access():
        print("⛔ DỪNG CHƯƠNG TRÌNH: hãy cấp quyền Editor cho service account rồi chạy lại.")
        return
    print("✅ Có quyền ghi Google Sheet.")

    try:
        rows = get_unposted_rows(limit=MAX_POSTS_PER_RUN)
    except Exception as e:
        print(f"❌ Không đọc được tab {RECRUIT_TAB_NAME} sau nhiều lần thử: {e}")
        print(
            "💡 Kiểm tra: tắt VPN/proxy, tạm tắt antivirus/firewall quét HTTPS, "
            "thử đổi mạng (vd: hotspot điện thoại) hoặc đổi DNS sang 8.8.8.8 / 1.1.1.1."
        )
        return

    if not rows:
        print("🎉 Không có bài nào cần đăng.")
        return
    print(f"📄 Tìm thấy {len(rows)} bài chưa đăng")

    disabled_accounts = recently_disabled_codes()
    newly_disabled = set()
    if disabled_accounts:
        print(f"🚫 Bỏ qua (đã phát hiện bị khóa gần đây): {', '.join(sorted(disabled_accounts))}")
    inactive_rows = {}
    summary = {"posted": 0, "reply_error": 0, "error": 0, "skipped": 0, "check": 0}

    for i, item in enumerate(rows):
        row_index = item["row_index"]
        data = item["data"]
        status = item.get("status", "")

        acc_code = str(data.get(COL_ACCOUNTS_CODE, "")).strip()
        job_content = data.get(COL_JOB_CONTENT, "").strip()
        thread_content = data.get(COL_THREAD_CONTENT, "").strip()
        topic = str(data.get(COL_TOPIC, "")).strip()
        image_url = data.get(COL_IMAGE, "").strip()

        print("=" * 60)
        print(f"📌 BÀI ĐĂNG {i+1}/{len(rows)} - ACCOUNT: {acc_code}")
        print(f"📍 ROW INDEX: {row_index}")

        if status == STATUS_CHECK:
            print(
                f"🔎 Dòng {row_index} đang ở trạng thái {STATUS_CHECK} (chưa xác minh bài có lên không) "
                "→ SKIP. Kiểm tra tài khoản thủ công rồi sửa cột Posted."
            )
            summary["skipped"] += 1
            continue

        if status == STATUS_YES_NO_LINK:
            if YES_WITHOUT_LINK_ACTION == "check":
                print(
                    f"🔎 Dòng {row_index}: Posted=YES nhưng cột {COL_LINK_POST} trống "
                    f"→ đánh dấu {STATUS_CHECK}, kiểm tra tay."
                )
                mark_check(
                    row_index,
                    f"⚠ Posted=YES nhưng thiếu {COL_LINK_POST}: cần kiểm tra thủ công",
                )
                summary["check"] += 1
                continue
            print(
                f"⚠ Dòng {row_index}: Posted=YES nhưng cột {COL_LINK_POST} trống "
                "→ xử lý như lỗi (ERROR), đăng lại."
            )
            status = ""

        if not acc_code or acc_code not in accounts_dict:
            print(f"⚠ Mã account '{acc_code}' không hợp lệ hoặc trống → SKIP")
            summary["skipped"] += 1
            continue

        if not accounts_dict[acc_code].get("active", True):
            print(
                f"⏭ Tài khoản {acc_code} đang tắt (cột {COL_ACC_ACTIVE}) "
                "→ SKIP (không ghi gì lên sheet)"
            )
            inactive_rows[acc_code] = inactive_rows.get(acc_code, 0) + 1
            summary["skipped"] += 1
            continue

        if acc_code.upper() in disabled_accounts:
            print(f"🚫 Tài khoản {acc_code} bị khóa → SKIP")
            if status not in (STATUS_REPLY_ERROR, STATUS_ACCOUNT_DISABLED.upper()):
                mark_account_disabled_row(row_index, acc_code)
            inactive_rows[acc_code] = inactive_rows.get(acc_code, 0) + 1
            summary["skipped"] += 1
            continue

        existing_post_url = ""
        if status == STATUS_REPLY_ERROR:
            existing_post_url = get_existing_post_url(data)
            if not existing_post_url:
                print(
                    "⚠ Dòng REPLY_ERROR nhưng không có link bài hợp lệ ở cột Link post → SKIP"
                )
                summary["skipped"] += 1
                continue

        thread_content_normalized = normalize_threads_content(thread_content)
        job_content_normalized = normalize_threads_content(job_content)

        if not job_content_normalized:
            print("⚠ Job Content (Bài chính) trống → SKIP")
            summary["skipped"] += 1
            continue

        main_text, overflow_parts = split_main_and_overflow(
            job_content_normalized, topic
        )
        thread_parts = split_by_blocks(
            thread_content_normalized, THREADS_CHAR_LIMIT, THREADS_CHAR_LIMIT
        )
        reply_parts = build_reply_parts(overflow_parts, thread_content_normalized)

        print(
            f"🧩 Kế hoạch: bài chính {len(main_text)} ký tự"
            + (
                " | comment: " + ", ".join(str(len(p)) for p in reply_parts) + " ký tự"
                if reply_parts
                else " | không có comment"
            )
        )

        if overflow_parts:
            print(
                f"✂️ Bài chính dài {len(job_content_normalized)} ký tự → "
                f"giữ {len(main_text)} ký tự ở bài chính, chuyển {len(overflow_parts)} đoạn dư xuống comment."
            )
        if len(thread_parts) > 1:
            print(f"✂️ Thread Content dài → tách thành {len(thread_parts)} comment.")

        skip_wait = False
        acc_info = accounts_dict[acc_code]
        bot = ThreadsBot(
            account_code=acc_code,
            email=acc_info["email"],
            password=acc_info["password"],
            headless=HEADLESS,
            username=acc_info.get("username", ""),
        )

        image_path = None
        post_url = existing_post_url
        try:
            await bot.start()

            if not post_url:
                if image_url:
                    try:
                        image_path = download_image(image_url)
                    except Exception as e:
                        raise Exception(f"❌ Tải ảnh thất bại: {e}")

                post_url = await bot.post(
                    text=main_text,
                    image_path=image_path,
                    topic=topic,
                )
                print(f"🔗 Post URL (Bài chính): {post_url}")
            else:
                print(f"🔁 Chỉ thử lại comment phụ cho bài: {post_url}")

            if reply_parts:
                try:
                    await bot.reply_to_post(
                        post_url=post_url, texts=reply_parts, main_text=main_text
                    )
                except AccountDisabledError:
                    raise
                except Exception as reply_err:
                    print(f"❌ Comment phụ thất bại: {reply_err}")
                    mark_reply_error(row_index=row_index, post_url=post_url)
                    summary["reply_error"] += 1
                else:
                    mark_posted(row_index=row_index, post_url=post_url)
                    summary["posted"] += 1
            else:
                mark_posted(row_index=row_index, post_url=post_url)
                summary["posted"] += 1

            print(f"✅ Đã xử lý dòng {row_index} cho {acc_code}")

        except AccountDisabledError as e:
            print(f"🚫 {e} → ghi '{STATUS_ACCOUNT_DISABLED}', sang dòng kế tiếp")
            disabled_accounts.add(acc_code.upper())
            newly_disabled.add(acc_code)
            remember_disabled(acc_code)
            inactive_rows[acc_code] = inactive_rows.get(acc_code, 0) + 1
            summary["skipped"] += 1
            skip_wait = True
            if post_url and not existing_post_url:
                mark_reply_error(row_index=row_index, post_url=post_url)
            elif status != STATUS_REPLY_ERROR:
                mark_account_disabled_row(row_index, acc_code)

        except PostUnverifiedError as e:
            print(f"🔎 {e}")
            mark_check(row_index, f"⚠ Cần kiểm tra thủ công: {str(e)[:180]}")
            summary["check"] += 1

        except Exception as e:
            print(f"❌ LỖI ĐĂNG BÀI CHO {acc_code}")
            print(str(e))
            traceback.print_exc()
            if post_url and not existing_post_url:
                mark_reply_error(row_index=row_index, post_url=post_url)
            else:
                mark_error(row_index, f"❌ {str(e)[:200]}")
            summary["error"] += 1

        finally:
            if image_path and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                except Exception:
                    pass
            await bot.close()
            print(f"🛑 Đã đóng trình duyệt của {acc_code}")

        if i < len(rows) - 1 and not skip_wait:
            wait_time = random.randint(*DELAY_BETWEEN_POSTS)
            print(f"⏳ Đang nghỉ ngẫu nhiên {wait_time} giây để chống Spam...")
            await asyncio.sleep(wait_time)

    print(
        f"📊 Kết quả: đăng xong {summary['posted']} | comment lỗi {summary['reply_error']} "
        f"| cần kiểm tra tay {summary['check']} | lỗi {summary['error']} | bỏ qua {summary['skipped']}"
    )
    if newly_disabled:
        print(f"🚫 Tài khoản vừa phát hiện bị khóa: {', '.join(sorted(newly_disabled))}")
    if inactive_rows:
        detail = ", ".join(f"{c} ({n} dòng)" for c, n in sorted(inactive_rows.items()))
        print(
            f"⏭ Dòng bị bỏ qua vì tài khoản bị khóa/tắt: {detail}. "
            "Các dòng này đã ghi 'Account disabled'; đổi AccountsCode sang tài khoản khác để đăng lại."
        )
    if os.path.exists(UNSYNCED_FILE):
        print(f"📝 Có trạng thái chưa ghi được lên sheet, xem: {UNSYNCED_FILE}")
    print("🎯 HOÀN TẤT QUÉT & ĐĂNG TẤT CẢ CÁC BÀI!")


if __name__ == "__main__":
    asyncio.run(run())