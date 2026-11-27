import os
import subprocess

BASE_DIR = os.getcwd()


def get_chrome_path():
    # Tìm đường dẫn Chrome cài đặt trên máy Windows của bạn
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ]
    for path in paths:
        if os.path.exists(path):
            return path
    return None


def manual_login():
    print("=== TOOL ĐĂNG NHẬP THỦ CÔNG ===")
    account_code = input("Nhập mã AccountsCode muốn đăng nhập (VD: AH01): ").strip()

    if not account_code:
        print("Mã không được để trống!")
        return

    profile_dir = os.path.join(BASE_DIR, "profiles", account_code)
    os.makedirs(profile_dir, exist_ok=True)

    chrome_path = get_chrome_path()
    if not chrome_path:
        print(
            "❌ Không tìm thấy Google Chrome trên máy tính của bạn. Hãy cài đặt Google Chrome!"
        )
        return

    print(f"\n✅ Đang mở Google Chrome thật cho tài khoản {account_code}...")

    # Mở Chrome hoàn toàn độc lập, tách biệt khỏi Playwright (trỏ thẳng tới threads.com)
    cmd = f'"{chrome_path}" --user-data-dir="{profile_dir}" "https://www.threads.com/login"'
    subprocess.Popen(cmd, shell=True)

    print("👉 Hãy tự đăng nhập (bằng Instagram hoặc Username đều được).")
    print(
        "👉 Điền mã xác nhận thoải mái, nút Continue sẽ sáng lên vì đây là trình duyệt thật 100%."
    )
    print("👉 Tắt trình duyệt VÀ quay lại Terminal này, NHẤN ENTER để hoàn tất.")
    input()

    print(f"✅ Đã lưu phiên đăng nhập thành công vào thư mục: profiles/{account_code}")


if __name__ == "__main__":
    manual_login()