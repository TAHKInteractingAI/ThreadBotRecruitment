import os
from playwright.sync_api import sync_playwright

BASE_DIR = os.getcwd()


def extract_cookies():
    print("=== TOOL TRÍCH XUẤT COOKIE (STORAGE STATE) ===")
    account_code = input("Nhập mã AccountsCode muốn lấy Cookie (VD: AH01): ").strip()

    if not account_code:
        print("Mã không được để trống!")
        return

    profile_dir = os.path.join(BASE_DIR, "profiles", account_code)

    if not os.path.exists(profile_dir):
        print(f"❌ Không tìm thấy thư mục: {profile_dir}")
        print("👉 Vui lòng chạy file login.py để đăng nhập tay trước nhé!")
        return

    # Tạo thư mục chứa cookies
    cookie_dir = os.path.join(BASE_DIR, "cookies")
    os.makedirs(cookie_dir, exist_ok=True)

    cookie_file = os.path.join(cookie_dir, f"{account_code}.json")

    with sync_playwright() as p:
        print(f"⏳ Đang trích xuất Cookie từ thư mục {profile_dir}...")

        # Mở lại profile bằng Playwright (ở chế độ chạy ngầm)
        ctx = p.chromium.launch_persistent_context(
            profile_dir,
            headless=True,
            channel="chrome",
            args=["--disable-blink-features=AutomationControlled"],
        )

        try:
            # Mở threads.com để đảm bảo cookie/localStorage của domain .threads.com được đồng bộ đầy đủ
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(
                "https://www.threads.com",
                wait_until="domcontentloaded",
                timeout=30000,
            )
            page.wait_for_timeout(4000)

            # Trích toàn bộ trạng thái đăng nhập ra file JSON
            ctx.storage_state(path=cookie_file)
        finally:
            ctx.close()

        print(f"✅ THÀNH CÔNG! Đã xuất file Cookie siêu nhẹ tại: {cookie_file}")

        # In dung lượng file
        size_kb = os.path.getsize(cookie_file) / 1024
        print(f"👉 Kích thước file hiện tại: {size_kb:.2f} KB")


if __name__ == "__main__":
    extract_cookies()