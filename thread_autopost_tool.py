import os
import sys
import time
import random
import traceback
import asyncio
import requests
import regex as re
import gspread
import pytz

from datetime import datetime
from PIL import Image
from oauth2client.service_account import ServiceAccountCredentials
from playwright.async_api import async_playwright

# Windows async event loop fix
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# =========================
# CONFIGURATION
# =========================
BASE_DIR = os.getcwd()
CREDENTIAL_FILE = "credentials.json"
RECRUIT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1YZrOO7Wb1fSCKeLNQnZMbXfNJuc7-kIz5ub1aXmzZdg/edit?usp=sharing"

# ===== TÊN TAB =====
RECRUIT_TAB_NAME = "Recruitment"
ACCOUNT_TAB_NAME = "Accounts"

# ===== TÊN CỘT TAB RECRUITMENT =====
COL_POSITION = "Position"
COL_JOB_CONTENT = "Job Content"
COL_THREAD_CONTENT = "Thread Content"
COL_TOPIC = "Topic"
COL_IMAGE = "Image URL"
COL_POSTED = "Posted"
COL_LINK_POST = "Link post"
COL_DATE = "Date"
COL_ACCOUNTS_CODE = "AccountsCode"

MAX_POSTS_PER_RUN = 999
THREADS_URL = "https://www.threads.net"

POST_DELAY_RANGE = (3, 6)
AFTER_POST_DELAY = (5, 8)
DELAY_BETWEEN_POSTS = (60, 180)


# =========================
# ULTILS
# =========================
def normalize_threads_content(text: str) -> str:
    if not text:
        return ""
    t = text.strip()
    t = text.replace("\r\n", "\n")
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


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
def connect_sheet(tab_name):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE, scope)
    client = gspread.authorize(creds)
    return client.open_by_url(RECRUIT_SHEET_URL).worksheet(tab_name)


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
            accounts[code] = {
                "email": str(row.get("Email", "")).strip(),
                "password": str(row.get("Password", "")).strip(),
            }
    return accounts


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
        if posted == "YES":
            continue
        results.append({"row_index": idx, "data": row})
        if len(results) >= limit:
            break
    return results


def mark_posted(row_index: int, post_url: str):
    sheet = connect_sheet(RECRUIT_TAB_NAME)
    tz_vn = pytz.timezone("Asia/Ho_Chi_Minh")
    now_time = datetime.now(tz_vn).strftime("%Y-%m-%d %H:%M:%S")
    sheet.update_cell(row_index, _col_index(RECRUIT_TAB_NAME, COL_POSTED), "YES")
    sheet.update_cell(row_index, _col_index(RECRUIT_TAB_NAME, COL_LINK_POST), post_url)
    sheet.update_cell(row_index, _col_index(RECRUIT_TAB_NAME, COL_DATE), now_time)


def mark_error(row_index: int, error_msg: str):
    sheet = connect_sheet(RECRUIT_TAB_NAME)
    tz_vn = pytz.timezone("Asia/Ho_Chi_Minh")
    now_time = datetime.now(tz_vn).strftime("%Y-%m-%d %H:%M:%S")
    sheet.update_cell(row_index, _col_index(RECRUIT_TAB_NAME, COL_POSTED), "ERROR")
    sheet.update_cell(row_index, _col_index(RECRUIT_TAB_NAME, COL_LINK_POST), error_msg)
    sheet.update_cell(row_index, _col_index(RECRUIT_TAB_NAME, COL_DATE), now_time)


def _col_index(tab_name: str, col_name: str) -> int:
    sheet = connect_sheet(tab_name)
    headers = sheet.row_values(1)
    for i, h in enumerate(headers, start=1):
        if h.strip() == col_name:
            return i
    raise Exception(f"❌ Không tìm thấy cột: {col_name} trong tab {tab_name}")


# ==========================================
# THREADS BOT (PLAYWRIGHT)
# ==========================================
class ThreadsBot:
    def __init__(self, account_code: str, email: str, password: str, headless=True):
        self.headless = headless
        self.account_code = account_code
        self.email = email
        self.password = password

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

    async def start(self):
        self.pw = await async_playwright().__aenter__()
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
        self.page = await self.context.new_page()

        await self.page.goto(THREADS_URL, wait_until="networkidle")
        is_logged_in = await self._is_logged_in()

        if not is_logged_in:
            print(
                f"⚠ Cookie hỏng hoặc chưa có cho {self.account_code}. Tiến hành auto-login..."
            )
            await self._login()
        else:
            print(f"✅ Tài khoản {self.account_code} đăng nhập thành công bằng Cookie.")

    async def close(self):
        if self.context:
            await self.context.close()
        if self.browser:
            await self.browser.close()
        if self.pw:
            await self.pw.stop()

    async def _is_logged_in(self) -> bool:
        try:
            await self.page.wait_for_timeout(3000)
            guest_1 = self.page.locator("text='Đăng nhập hoặc đăng ký'").first
            guest_2 = self.page.locator("text='Đăng nhập bằng tên người dùng'").first
            guest_3 = self.page.locator("text='Log in with phone'").first

            if (
                await guest_1.is_visible()
                or await guest_2.is_visible()
                or await guest_3.is_visible()
            ):
                return False
            return True
        except:
            return False

    async def _login(self):
        await self.page.goto(f"{THREADS_URL}/login", wait_until="networkidle")
        try:
            await self.page.wait_for_timeout(3000)
            try:
                switch_btn = self.page.locator(
                    "text='Đăng nhập bằng tên người dùng', text='Log in with phone', text='Log in with email'"
                ).first
                if await switch_btn.is_visible(timeout=5000):
                    print("🔄 Đang click mở Form gõ Email/Password...")
                    await switch_btn.click(force=True)
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
            guest_check = self.page.locator(
                "text='Đăng nhập bằng tên người dùng', button:has-text('Đăng nhập'), button:has-text('Log in')"
            ).first
            if await guest_check.is_visible():
                raise Exception(
                    "Kẹt ở màn hình đăng nhập (Bị vướng Captcha hoặc Checkpoint). Hãy dùng script get_cookie trên máy tính."
                )

            print(f"✅ Đăng nhập xong cho {self.account_code}!")
            await self.context.storage_state(path=self.cookie_file)
            print(f"💾 Đã lưu Cookie mới vào: {self.cookie_file}")

        except Exception as e:
            await self.page.screenshot(
                path=os.path.join(
                    self.error_dir, f"error_login_{self.account_code}.png"
                )
            )
            raise Exception(f"❌ Login tự động thất bại. Lỗi: {e}")

    async def _get_recent_post_urls(self) -> set:
        username = await self.get_profile_name()
        if not username:
            return set()
        await self.page.goto(f"{THREADS_URL}/{username}", wait_until="networkidle")
        await self.page.wait_for_timeout(3000)

        urls = set()
        link_locators = self.page.locator("a[href*='/post/']")
        count = await link_locators.count()
        for i in range(min(5, count)):
            href = await link_locators.nth(i).get_attribute("href")
            if href:
                full_url = f"{THREADS_URL}{href}" if href.startswith("/") else href
                urls.add(full_url)
        return urls

    async def _wait_for_new_post(self, old_urls: set) -> str:
        username = await self.get_profile_name()
        if not username:
            return ""

        for attempt in range(6):
            print(f"   ⏳ F5 tải lại trang cá nhân (Lần {attempt + 1}/6)...")
            await self.page.goto(f"{THREADS_URL}/{username}", wait_until="networkidle")
            await self.page.wait_for_timeout(5000)

            link_locators = self.page.locator("a[href*='/post/']")
            count = await link_locators.count()
            for i in range(min(5, count)):
                href = await link_locators.nth(i).get_attribute("href")
                if href:
                    full_url = f"{THREADS_URL}{href}" if href.startswith("/") else href
                    if full_url not in old_urls:
                        return full_url

            await self.page.wait_for_timeout(3000)

        await self.page.screenshot(
            path=os.path.join(
                self.error_dir, f"error_missing_post_{self.account_code}.png"
            )
        )
        return ""

    async def post(self, text: str, image_path: str | None = None, topic: str = ""):
        if not text.strip():
            raise ValueError("❌ Nội dung bài post trống")

        print("🔍 Đang ghi nhận danh sách bài cũ trên tường để đối chiếu...")
        old_post_urls = await self._get_recent_post_urls()

        await self._open_composer()
        await self._type_text(text)

        if topic:
            clean_topic = topic.replace("#", "").strip()
            if clean_topic:
                print(f"🏷️ Đang gắn Chủ đề (Topic): {clean_topic}...")
                await self.page.keyboard.type(f"\n\n#{clean_topic}", delay=150)
                await self.page.wait_for_timeout(2000)
                await self.page.keyboard.press("Enter")
                await self.page.wait_for_timeout(1000)

        time.sleep(2)

        if image_path:
            await self._upload_image(image_path)

        print("🚀 Đang bấm nút đăng bài...")
        await self._submit_post()

        print("🔍 Đang chờ Threads load xong bài viết mới...")
        post_url = await self._wait_for_new_post(old_post_urls)
        if not post_url:
            raise Exception(
                "❌ Post KHÔNG xuất hiện trên Threads profile. (Đã lưu ảnh error_missing_post để kiểm tra)"
            )
        return post_url

    async def reply_to_post(self, post_url: str, text: str):
        if not text.strip():
            return

        print("💬 Đang tiến hành thả comment phụ (Thread Content)...")
        await self.page.goto(post_url, wait_until="networkidle")
        await self.page.wait_for_timeout(random.randint(4000, 6000))

        try:
            reply_btn = self.page.locator(
                "svg[aria-label='Trả lời'], svg[aria-label='Reply']"
            ).first
            await reply_btn.wait_for(state="visible", timeout=10000)
            await reply_btn.click(force=True)  # Ép click nút trả lời
            await self.page.wait_for_timeout(2000)

            editor = self.page.locator("div[contenteditable='true']").last
            await editor.wait_for(state="visible", timeout=10000)
            await editor.click(force=True)

            await self._type_text(text)
            await self.page.wait_for_timeout(1000)

            await self._submit_post()
            print("✅ Đã đăng comment phụ thành công!")
            await self.page.wait_for_timeout(4000)
        except Exception as e:
            await self.page.screenshot(
                path=os.path.join(
                    self.error_dir, f"error_reply_{self.account_code}.png"
                )
            )
            print(f"⚠ Không thể comment phụ. Lỗi: {e}")

    async def get_profile_name(self) -> str:
        try:
            el = await self.page.wait_for_selector("a[href^='/@']", timeout=5000)
            return await el.get_attribute("href")
        except:
            return ""

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
            except:
                try:
                    trigger_vi = self.page.locator("text='Có gì mới?'").first
                    await trigger_vi.wait_for(state="visible", timeout=3000)
                    await trigger_vi.click(force=True)
                except:
                    try:
                        trigger_en = self.page.locator('text="What\'s new?"').first
                        await trigger_en.wait_for(state="visible", timeout=3000)
                        await trigger_en.click(force=True)
                    except:
                        plus_btn = self.page.locator(
                            "svg[aria-label='Tạo'], svg[aria-label='Create'], svg[aria-label='Bắt đầu thread mới']"
                        ).first
                        await plus_btn.click(force=True, timeout=5000)

            time.sleep(2)
            editor = self.page.locator("div[contenteditable='true']").first
            await editor.wait_for(state="visible", timeout=15000)
            await editor.click(force=True)
            time.sleep(random.uniform(*POST_DELAY_RANGE))

        except Exception as e:
            await self.page.screenshot(
                path=os.path.join(
                    self.error_dir, f"error_open_composer_{self.account_code}.png"
                )
            )
            raise Exception(f"Không thể mở khung đăng bài: {e}")

    async def _type_text(self, text: str):
        delay_typing = random.randint(15, 30)
        await self.page.keyboard.type(text, delay=delay_typing)
        time.sleep(random.uniform(*POST_DELAY_RANGE))

    async def _upload_image(self, image_path: str):
        file_input = self.page.locator("input[type='file']").first
        await file_input.set_input_files(image_path)
        await self.page.wait_for_timeout(5000)

    async def _submit_post(self):
        try:
            # Lấy ô soạn thảo cuối cùng
            editor = self.page.locator("div[contenteditable='true']").last
            await editor.click(force=True, timeout=2000)
            await self.page.wait_for_timeout(500)
        except:
            pass

        try:
            post_btn = (
                self.page.locator("div[role='button']")
                .filter(has_text=re.compile(r"^(Đăng|Post)$"))
                .first
            )
            await post_btn.scroll_into_view_if_needed(timeout=2000)
            await post_btn.click(force=True, timeout=3000)
        except Exception:
            await self.page.keyboard.down("Control")
            await self.page.keyboard.press("Enter")
            await self.page.keyboard.up("Control")

        time.sleep(8)


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

    rows = get_unposted_rows(limit=MAX_POSTS_PER_RUN)
    if not rows:
        print("🎉 Không có bài nào cần đăng.")
        return
    print(f"📄 Tìm thấy {len(rows)} bài chưa đăng")

    for i, item in enumerate(rows):
        row_index = item["row_index"]
        data = item["data"]

        acc_code = str(data.get(COL_ACCOUNTS_CODE, "")).strip()
        job_content = data.get(COL_JOB_CONTENT, "").strip()
        thread_content = data.get(COL_THREAD_CONTENT, "").strip()
        topic = str(data.get(COL_TOPIC, "")).strip()
        image_url = data.get(COL_IMAGE, "").strip()

        print("=" * 60)
        print(f"📌 BÀI ĐĂNG {i+1}/{len(rows)} - ACCOUNT: {acc_code}")
        print(f"📍 ROW INDEX: {row_index}")

        if not acc_code or acc_code not in accounts_dict:
            print(f"⚠ Mã account '{acc_code}' không hợp lệ hoặc trống → SKIP")
            continue

        if not job_content:
            print("⚠ Job Content (Bài chính) trống → SKIP")
            continue

        job_content_normalized = normalize_threads_content(job_content)
        total_length = len(job_content_normalized)

        if topic:
            total_length += len(topic.replace("#", "").strip()) + 3

        if total_length > 500:
            error_msg = f"❌ Lỗi: Bài chính quá dài ({total_length}/500 ký tự). Vui lòng cắt bớt chữ!"
            print(f"🛑 BỎ QUA BÀI ĐĂNG: {error_msg}")
            mark_error(row_index, error_msg)
            continue

        thread_content_normalized = normalize_threads_content(thread_content)
        if len(thread_content_normalized) > 500:
            error_msg = f"❌ Lỗi: Comment phụ quá dài ({len(thread_content_normalized)}/500 ký tự). Vui lòng cắt bớt chữ!"
            print(f"🛑 BỎ QUA BÀI ĐĂNG: {error_msg}")
            mark_error(row_index, error_msg)
            continue

        acc_info = accounts_dict[acc_code]
        bot = ThreadsBot(
            account_code=acc_code,
            email=acc_info["email"],
            password=acc_info["password"],
            headless=True,  # Nhớ để True khi up lên Github
        )

        image_path = None
        try:
            await bot.start()

            if image_url:
                try:
                    image_path = download_image(image_url)
                except Exception as e:
                    raise Exception(f"❌ Tải ảnh thất bại: {e}")

            post_url = await bot.post(
                text=job_content_normalized,
                image_path=image_path,
                topic=topic,
            )
            print(f"🔗 Post URL (Bài chính): {post_url}")

            if thread_content_normalized:
                await bot.reply_to_post(
                    post_url=post_url, text=thread_content_normalized
                )

            mark_posted(row_index=row_index, post_url=post_url)

            if image_path and os.path.exists(image_path):
                os.remove(image_path)

            print(f"✅ Đã xử lý (Đăng + Comment) cho {acc_code}")

        except Exception as e:
            print(f"❌ LỖI ĐĂNG BÀI CHO {acc_code}")
            print(str(e))
            traceback.print_exc()

        finally:
            await bot.close()
            print(f"🛑 Đã đóng trình duyệt của {acc_code}")

        if i < len(rows) - 1:
            wait_time = random.randint(*DELAY_BETWEEN_POSTS)
            print(f"⏳ Đang nghỉ ngẫu nhiên {wait_time} giây để chống Spam...")
            time.sleep(wait_time)

    print("🎯 HOÀN TẤT QUÉT & ĐĂNG TẤT CẢ CÁC BÀI!")


if __name__ == "__main__":
    asyncio.run(run())
